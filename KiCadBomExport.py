#!/usr/bin/env python2
####################################################
#
#  KiCad BOM Converter and Coster
#  ------------------------------
#
#   PYTHON 2.74 (KiCad does not seem to support Python 3)
#
#  
#   Processes the KiCad intermediate BOM file, allowing you to:
#   - Export to CSV / XML
#   - Group common components into a single row
#   - Cost the BOM using online tool such as SupplyFrame's FindChip.com
#
#   Command Line Arguments:
#   -h / help   Display Help message
#   -g / group  Group like components into a single row
#   -i / input  Input file (in KiCad's intermediate XML format)
#   -o / output Output file name (will be appended with .CSV and .XML)
#   -f          Use the FindChips API
#   -a / apikey The API Key from the online costing service
#
#   Required Template Fieldname (within EESchema, Preferences / Schematic Editor Options / Template Field Name)
#   - Mfg_Part_No   Store the manufacturer's part number.  Be specific down to package if you use the online costing, to get specific part
#
#   Suggested Template Fieldnames (not used here, but I found them useful):
#   - Manufacturer
#   - Alt_Manufacturer
#   - Alt_Mfg_Part_No
#   - Contract_Manuf_SKU
#
#   USAGE in KiCad (As at Jun 2015):
#       - Open EESchema
#       - TOOLS menu / Generate Bill of Materials
#       - Add a new item with a cmd line similar to:
#           python "<<path to python script>>KiCadBomExport.py" -i "%I" <<other cmd line params>>
#       (suggest you don't use the "%O" output file parameter in order to prevent intermediate netlist file being overwritten)
#
#   Troubleshooting:
#   - A log file is created in user home directory
#
#   Further Enhancements Identified:
#   - Add command-line option to cost based on number of boards (currently only assumes 1)
#   - Use a config file to restrict distributors to s preferred set
#
#  Original Version:
#   Andrew Retallack    andrewr(at)crash-bang(dot)com
#   June 2015
#   www.crash-bang.com
#
#   License: GNU General Public License v3.0
#       http://www.gnu.org/licenses/gpl-3.0.html







import csv
import sys
import os
import getopt
import httplib
import urllib
import json
import logging
from os.path import expanduser
from xml.etree.ElementTree import Element, SubElement, ElementTree



#List of initial field names for our CSV file - is expanded as we go.
CSVFieldNames = ['Reference','Value','Footprint','Count','Datasheet']

#These are the XML Tags used in the KiCad intermediate netlist file.  Declared here for ease of maintenance
tagComponent = 'components'
tagRef = 'ref'
tagValue = 'value'
tagFootprint = 'footprint'
tagFields = 'fields'
tagDatasheet = 'datasheet'

fldMfgPartNo = 'Mfg_Part_No'

#If we use an online pricing service (FindChips)
pricingService = ''   #The online Pricing service we're using (F = FindChips)
apiKey = '' #The online service's API Key

listOutput = [] #Stores components once processed
curPart = {} #elements of the current part being processed

#Create a Log File in the user home directory
logFile = os.path.join(expanduser("~"), 'Kicad_BOM_Logging.txt')
logging.basicConfig(filename=logFile, filemode='w', level = 'INFO', format='%(asctime)-15s %(levelname)-10s %(message)s')
logger = logging.getLogger('Standard')
logger.info('\n----------------------------\n')
logger.info('Script Started')


def main(argv):
    global pricingService
    global apiKey
    #Default values - will change these
    fileIn = ''
    fileOut = ''
    groupParts = False
    
    pricingService = '' #Which pricing service to use (currently 'F' = FindChips)
    
#####
# COMMENT OUT FOR COMMAND LINE - THIS IS HERE FOR DEVELOPMENT TESTING ONLY
#####
#    fileIn = ''  #replace with filename
#    groupParts = True
#    apiKey = ''    #Put your API key here
#    pricingService= 'F'
#    fileOut = ''   #replace with filename
######



    ########################################################
    #Process the Command Line
    logger.info('Command Line Arguments: %s',str(argv))
        
    try:
        opts, args = getopt.getopt(argv, "hgfa:i:o:", ["help", "group", "apikey=", "input=", "output="])
    except getopt.GetoptError:
        printUsage()
        logger.error('Invalid argument list provided')
        sys.exit(2)
        
    for opt, arg in opts:
        if opt in ("-i", "--input"):
            fileIn = arg

        if opt in ("-o", "--output"):
            fileOut = arg

        if opt in ("-a", "--apikey"):
            apiKey = arg

        if opt in ("-f", "--findchips"):
            pricingService = opt

        if opt in ("-g", "--group"):
            groupParts = True
            
        if opt in ("-h", "--help"):
            printUsage()

    #Perform error checking to make sure cmd-line params passed, files exist etc.
    checkParams(fileIn)


    #If no output file specified, create one:
    if fileOut == '':
        fileOut = fileIn[:fileIn.rfind('.')] + '_BOM'    #strip the extension off the Input file
        logger.info('No OUTPUT Filename specified - using: %s', fileOut)

    
    ########################################################
    #Parse and process the XML netlist
    eTree = ElementTree()
    eTree.parse(fileIn)
    root = eTree.getroot()
    section = root.find(tagComponent)  #Find the component-level in the XML Netlist

    #Loop through each component
    for component in section:
        processComponent(listOutput, component, groupParts) #Extract compoonent info and add to listOutput 

    #Get Pricing if we need to.  Need a Manufacturer Part No, and a Pricing Service to be specified
    if fldMfgPartNo in CSVFieldNames and pricingService != '':
        getPricing(listOutput)


    ########################################################
    #Generate the Output CSV File
    with open(fileOut+'.csv', 'wb') as fOut:
        csvWriter = csv.DictWriter(fOut, delimiter = ',', fieldnames = CSVFieldNames)
        csvWriter.writeheader()
        utf8Output = []
        for row in listOutput:
            utf8Row = {}
            for key in row:
                utf8Row[key] = row[key].encode('utf-8')
            utf8Output.append(utf8Row)
        csvWriter.writerows(utf8Output)

    #Print and Log the file creation
    print('Created CSV File')
    logger.info('Created CSV file with %i items: %s', len(listOutput), fileOut+'.csv')


    ########################################################
    #Generate the Output XML File

    parent = Element('schematic')   #Create top-level XML element

    #Loop through each Component and create a child of the XML top-level
    for listItem in listOutput:
        child = SubElement(parent, 'component')

        #Loop through each attribute of the component.
        #  We do it this way, in order to preserve order of elements - more logical output
        for key in CSVFieldNames:
            if key in listItem:
                attribute = SubElement(child, key.replace(' ','_').replace('&','_'))    #XML doesn't like spaces in element names, so replace them with "_"
                attribute.text = listItem[key]
        
    #Output to XML file
    ET = ElementTree(parent)
    ET.write(fileOut+'.xml')

    #Print and Log the file creation
    print('Created XML File')
    logger.info('Created XML file with %i items: %s', len(listOutput), fileOut+'.xml')

    #Close the log file - we're done!
    logging.shutdown()

    ########################################################
    #All Done
    ########################################################
    





########################################################
#
#  Routine to process a single component from KiCad XML Netlist
#   Adds component to listOutput
#   Extracts all component fields/attributes
#
#   PARAMS: listOutput      list to add processed component to
#           xmlComponent    component to process
#           groupParts      should same parts be grouped into single row
#
########################################################

def processComponent(listOutput, xmlComponent, groupParts):

    bDup = False    #Is this part a duplicate of an already-existing one?
    curPart = {}    #Dict of current part's attributes

    #Extract the component name
    curPart['Reference'] = xmlComponent.attrib[tagRef]

    #Component Value
    if xmlComponent.find(tagValue) != None: curPart['Value'] = xmlComponent.find(tagValue).text
    if xmlComponent.find(tagFootprint) != None: curPart['Footprint'] = xmlComponent.find(tagFootprint).text
    if xmlComponent.find(tagDatasheet) != None: curPart['Datasheet'] = xmlComponent.find(tagDatasheet).text
    curPart['Count'] = '1'
        

    #if there are additional Attributes/Fields, add them
    if xmlComponent.find(tagFields)!= None:
        processFields(xmlComponent.find(tagFields), curPart)

    #If user wants parts grouped, then check whether this is a duplicate component.
    #If it is, add this Reference to existing part.
    if groupParts == True:

        #Loop through existing already-processed items
        for listItem in listOutput:
            #First, de-dup by manufacturer part no (if one exists)
            if fldMfgPartNo in curPart and fldMfgPartNo in listItem:
                if listItem[fldMfgPartNo] == curPart[fldMfgPartNo] and curPart[fldMfgPartNo] != '-':
                    bDup = True
                    listItem['Reference'] = listItem['Reference'] + ';' + curPart['Reference']  #Add current part reference to the existing part
                    if listItem['Value'] != curPart['Value']: 
                        listItem['Value'] = listItem['Value'] + ';' + curPart['Value']  #Add current part value to the existing part
                    listItem['Count'] = str(int(listItem['Count']) + 1) #Increase the count of times this part is used
                    break

            #No Manufac Part No - so de-dup on Value and Footprint
            elif listItem['Value'] == curPart['Value'] and listItem['Footprint'] == curPart['Footprint']:
                bDup = True
                listItem['Reference'] = listItem['Reference'] + ';' + curPart['Reference']  #Add current part reference to the existing part
                listItem['Value'] = listItem['Value']
                listItem['Count'] = str(int(listItem['Count']) + 1) #Increase the count of times this part is used
                #
                #  Future Development: merge in other fields here as well
                #
                break

    #If no duplicates found, then add this as a new component
    if bDup == False:
        listOutput.append(curPart)


    
########################################################
#
#  Routine to process fields for a specific component:
#   Adds any new fields to the master field list
#   Adds field values to current part
#
#   PARAMS: xmlFields   fields from the component in the XML netlist
#           curPart     dictionary of fields for the current part
#
########################################################

def processFields(xmlFields, curPart):
    tagField = 'name'
    fieldName = ''
    fieldValue=''
    
    #Loop through all fields for this component in the XML netlist
    for field in xmlFields:
        fieldName = field.attrib[tagField]  #get the field name
        fieldValue = field.text     #Get the field value
        curPart[fieldName] = fieldValue #Add to the current part

        #Check if this field name is in the Master list.  If not, add it
        if not(fieldName in CSVFieldNames): CSVFieldNames.append(fieldName)
    


########################################################
#
#  Routine to get pricing for parts from online service
#   Requests the pricing
#   Parses and adds volume-pricing
#
#   PARAMS: listOutput  the list of components to retrieve pricing for.
#                       Is also updated with the prices once retrieved
#
########################################################

def getPricing(listOutput):

    global CSVFieldNames

    #Connect to FindChips
    con = httplib.HTTPConnection("api.findchips.com")

    #For each component in the list
    for listItem in listOutput:

        #Only retrieve prices if a Manufacturer Part Number exists
        if fldMfgPartNo in listItem:
            #code up the request
            params = urllib.urlencode({'apiKey': apiKey, 'part': listItem[fldMfgPartNo], 'limit':15, 'hostedOnly':'false', 'authorizedOnly':'false', 'exactMatch':'true'})
            #Sent the request
            con.request("GET", "/v1/search?"+params)
            #Retrieve result
            r = con.getresponse()

            #If return status was not "OK" then log the error
            if r.status != 200:
                logger.error('Error returned by online pricing engine: %s', r.reason)
                logger.info('params: %s', params)
                print("Error returned by pricing service - check log file")
                return
            else:
                logger.info('Online Pricing: Retrieved for %s', listItem[fldMfgPartNo])
                logger.info('params: %s', params)


            #Now process the JSON response received from the online service
            jdata = json.loads(r.read().decode('utf-8'))
            #Loop through each supplier
            #
            #  FUTURE ENHANCEMENT: Compare this to a user-specified list of suppliers, and only process those
            #
            for supplier in jdata['response']:
                #Loop through each part/volume/price returned by the supplier
                for part in supplier['parts']:
                    lastQty = 0
                    lastPrice = 0
                    lastCurrency = ''
                    #Now we (clumsily) loop through all price-volume breaks, and find the smallest volume break that meets our requirements
                    if 'distributorItemNo' in part: lastPartNo = part['distributorItemNo']
                    for price in part['price']:
                        if price['quantity'] > int(listItem['Count']) and lastQty > 0:
                            break
                        lastQty = price['quantity']
                        lastPrice = price['price']
                        lastCurrency = price['currency']

                    #If we have found a price/volume record, add it to the current part.
                    if lastQty > 0:
                        listItem[str(supplier['distributor']['name'])+'_PARTNO'] = str(lastPartNo)  #Distributor Part No
                        listItem[str(supplier['distributor']['name'])+'_QTY'] = str(lastQty)        #Quantity
                        listItem[str(supplier['distributor']['name'])+'_CURRENCY'] = str(lastCurrency)  #Currency
                        listItem[str(supplier['distributor']['name'])+'_PRICE'] = str(lastPrice)    #Price
                        #Also add this supplier to the Master Field list if it doesn't already exist
                        if not(str(supplier['distributor']['name'])+'_PARTNO' in CSVFieldNames): CSVFieldNames.append(str(supplier['distributor']['name'])+'_PARTNO')
                        if not(str(supplier['distributor']['name'])+'_QTY' in CSVFieldNames): CSVFieldNames.append(str(supplier['distributor']['name'])+'_QTY')
                        if not(str(supplier['distributor']['name'])+'_PRICE' in CSVFieldNames): CSVFieldNames.append(str(supplier['distributor']['name'])+'_PRICE')
                        if not(str(supplier['distributor']['name'])+'_CURRENCY' in CSVFieldNames): CSVFieldNames.append(str(supplier['distributor']['name'])+'_CURRENCY')
                        
            #Close the connection with the online service
            con.close()


########################################################
#
#  Routine to print command-line help
#
########################################################

def printUsage():
    print("Export a KiCad netlist to CSV and XML formats")
    print("---------------------------------------------")
    print("usage: -i <Input FName> -o <Output FName> -a <API Key> -f -d")
    print("\n-f:\tUse the FindChips online pricing engine")
    print("-g:\tGroup duplicate parts into a single entry")
    print("<Input FName>:\tPath and filename of netlist")
    print("<Output FName>:\tOutput path and filename (no extension) for XML ans CSV")
    print("<API Key>:\tAPI Key for online pricing engine (FindChips / Octopart)")
    
    
    

########################################################
#
#  Routine to check command line parameters
#   Does input file exist
#   If online service specified, is there an API Key
#
#   PARAMS: fileIn  Name of input file
#
########################################################

def checkParams(fileIn):

    #Ensure filenames provided
    if fileIn == '':
        logger.error('No INPUT Filename was specified')
        printUsage()
        print("\n***ERROR*** Input Filename missing");
        sys.exit(2)

    if os.path.isfile(fileIn) == False:
        logger.error('The INPUT file does not exist: %s', fileIn) 
        printUsage()
        print("\n***ERROR*** Input File does not exist: " + fileIn)
        sys.exit(2)


    #If an online pricing service being used, check if it needs an API Key
    if (pricingService == 'F' or pricingService == 'O') and len(apiKey) == 0:
        logger.error('Have asked for online pricing %s but no API Key given.', pricingService) 
        printUsage()
        print('\n***ERROR*** Have asked for online pricing %s but no API Key given.', pricingService)
        sys.exit(2)
        

if __name__ == "__main__":
    main(sys.argv[1:])
    

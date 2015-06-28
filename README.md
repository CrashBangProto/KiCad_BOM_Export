# KiCad_BOM_Export
KiCad Add-in Script to export and cost a BOM
Written in Python 2.74, as KiCad seems not to support Python 3

This is a simple script process the KiCad intermediate netlist file, and produce a BOM in both CSV and XML Format

###Features:
- Optionally groups same components into a single row (cmd line paramater "-g")
- Optionally costs the BOM using FindChips API (cmd line parameters "-f" and "-a [apikey]")
- Produces both CSV and XML output

###Command Line Parameters
   -h / help   Display Help message
   -g / group  Group like components into a single row
   -i / input  Input file (in KiCad's intermediate XML format)
   -o / output Output file name (will be appended with .CSV and .XML).  If none specified, output file is input file + "_BOM"
   -f          Use the FindChips API
   -a / apikey The API Key from the online costing service


###Usage from within KiCad:
  - Open EESchema
  - TOOLS menu / Generate Bill of Materials
  - Add a new item with a cmd line similar to:
    python "[path to python script]KiCadBomExport.py" "-i %I" [other cmd line params]
    
    eg: python "c:\KiCadBomExport.py" -i "%I" -g -f -a xxxxxxxx
    Suggest you don't use the standard "%O" output file parameter, or the .xml output will overwrite the intermediate Netlist file.

###Usage with SupplyFrame's FindChips API
You need to request an API from SupplyFrame.  Go to: http://dev.supplyframe.com/
You will be provided with an API key.  Use this in your command line call after the "-a" argument

Within KiCad EESchema, you will need to define a manufacturer part number:
  - within EESchema, Preferences menu / Schematic Editor Options / Template Field Name
  - Add "Mfg_Part_No" to store the manufacturer's part number.  Be specific down to package to avoid vague search results


###Using the output files
I find that Excel doesn't read the CSV file in well, as it tries to convert data types.
I prefer to use Excel to open the XML output file:
  - right-click on XML output file, then "open with Excel"
  - in Excel choose to open as an XML Table
  - Voila!


###Troubleshooting:
  - A log file is created in user home directory (eg. on Windows: C:\User\[your username]\Kicad_BOM_Logging.txt)

Further Enhancements Identified:
  - Add command-line option to cost based on number of boards (currently only assumes 1)
  - Use a config file to restrict distributors to a preferred set

Please Contribute!  I found this useful, but there is much more that the script can achieve!


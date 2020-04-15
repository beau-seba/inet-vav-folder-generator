# inet-vav-folder-generator
Reads an EBO export template and a number of Excel spreadsheets listing information about INET VAV's returned by a Search, and generates a single .xml file which can be imported to create many system folders including a graphic page, trends, and alarms for each one.

How to use:

	Import the Search into Workstation from the Export in the "./resources" folder.

	Execute the Search query to list all -DmprPos points in the INET Interface, then copy the whole list view (Ctrl + a, Ctrl + c) to an Excel spreadsheet and save it in the included "./input" folder, or somewhere convenient.

	Run the executable and first select the Export Template from the "./input" folder, then select the spreadsheet(s), and a new Export will be created in the "./output" folder.

	Import the generated .xml file to Workstation under "~/System Folder/VAVs/" and many folders will be created containing a graphic, trends, etc., and all should be properly bound, as long as all points use the standard INET naming convention for ASC's.

	Go through the created objects, and check for any binding errors indicated by a blue triangle in the Bindings Editor. ("~/System/Binding Diagnostics" may also be useful) 

	NOTE: Sometimes I find space sensors with names changed, and on a couple occasions I have seen dual duct VAVs, which don't fit this template.

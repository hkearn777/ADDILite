# ADDILite
Program Code Modeler

Prerequisites: dotnet 8 and PlantUML. Dotnet can be downloaded from Microsoft. PlantUML can be loaded in my Releases.

<img width="596" alt="image" src="https://github.com/hkearn777/ADDILite/assets/110694374/b98f1f6a-72a3-4d61-a0bb-7a1516265d56">

Before actually pressing the ADDILite button, which will start the modeling process be sure to create the directory structure and fill in some of the folders.

Here is the structue/folders needed.
/App-folder    -- name of the system/application ie., the ROCS system. Not ony subfolders go here but also the Data Gathering Form.xls (filled in)

/App-folder/OUTPUT    -- this holds all output created including ADDILite.xls, ADDILite Log.csv, puml, and temp files.

/App-folder/JOBS      -- this holds all the JCL jobs; this is where to put the *.jcl files.

/App-folder/SOURCES    -- this holds all code sources, such as COBOL, Easytrieve, sub-routines, and all copybooks, includes, macros

/App-folder/OUTPUT/DEBUG  -- holds any debugging if turned on-Optional!

/App-folder/OUTPUT/PUML    -- holds the puml files to be flowcharted-Optional!

/App-folder/OUTPUT/SVG      -- holds the SVG (flowcharts)-Optional!

Options to run: 
- Scan mode only - creates a very limited .xls spreadsheet.
- Log Stmt Array - create the internal working files by JCL, COBOL or Eastrieve statement (not by lines).
- Delimiter - Vertical bar incase you need to change the delimiter value.
- Click the ADDILite button to begin the process
- Click the Close button to end the program.


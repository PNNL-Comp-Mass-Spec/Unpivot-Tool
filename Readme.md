Unpivot Tool

== Overview ==
This program reads in a delimited text file that is in crosstab (aka PivotTable) 
format and writes out a new file where the data has been unpivotted.

== Syntax ==

UnpivotTool.exe /I:InputFilePath [/O:OutputFolderName]
 [/F:FixedColumnCount] [/C:ColumnSepChar] [/B] [/N]

The input file path can contain the wildcard character *.  
If a wildcard is present, then all matching files will be processed.

The output folder name is optional.  If omitted, the output files 
will be created in the same folder as the input file.  
If included, then a subfolder is created with the name OutputFolderName.

Use /F to define the number of fixed columns (default is /F:1).  
When unpivotting, data in these columns will be written to every row in the output file.

The default column separation character is the tab character.  .
Use /C to define an alternate character.  For example, use /C:, for a comma.
For a space, use /C:space

Use /B to skip writing blank column values to the output file.

Use /N to skip writing Null values to the output file (as indicated by the word 'null').

== Example Usage ==

UnpivotTable.exe /I:ExamplePivotTable.txt /F:2 /B /N

See ExamplePivotTable.txt for an example input file 
and ExamplePivotTable_Unpivot.txt for an example output file.

-------------------------------------------------------------------------------
Written by Matthew Monroe for the Department of Energy (PNNL, Richland, WA)
Copyright 2009, Battelle Memorial Institute.  All Rights Reserved.

E-mail: matthew.monroe@pnnl.gov or proteomics@pnnl.gov
Website: http://panomics.pnnl.gov/ or http://omics.pnl.gov
-------------------------------------------------------------------------------

Licensed under the Apache License, Version 2.0; you may not use this file except 
in compliance with the License.  You may obtain a copy of the License at 
http://www.apache.org/licenses/LICENSE-2.0

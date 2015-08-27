# IntelliPlugin
An Excel plugin written in VB.NET to assist the Data team in their data building process.

<h1>Goals</h1>
-Add 'pick up where you left off' capability to Population and Validation <br/>
---Able to check to what point was the csv read up to and continue the process from there <br/>
-Add 'Stop Process' capability to Population and Validation <br/>
---Currently can't stop the process without completely kill 'Excel' process via Task Manager <br/>
-Possibly revisit Validation and Population's write-to-file <br/>
---Currently store all of the data in memory via stringbuilder, then write all at once to file <br/>
---Keeping everything in memory might cause problem as it hogged resources and more prone to errors <br/>

|06/10/2015|
|----------|
|<b>Population - Performance</b>|
|-Function: Import CSV file -> Calculate Formulas -> Output CSV file|
|-Input File: 2138868 PNs|
|-Performance: 7935.967s @ 5000 (Write to file only after Export is filled)|

|06/08/2015|
|----------|
|<b>Population - Performance</b>|
|-Function: Import CSV file -> Calculate Formulas -> Output CSV file|
|-Input File: 2138868 PNs|
|-Performance: 8232.44s @ 5000 (Write to file only after Export is filled)|

|06/02/2015|
|----------|
|<b>Population - Performance</b>|
|-Function: Import CSV file -> Calculate Formulas -> Output CSV file|
|-Input File: 2138868 PNs|
|-Performance: 7916.559s @ 5000 (Write to file only after Export is filled)|
|<b>Validator - Performance</b>|
|-Function: Import CSV file -> Calculate Formulas & Determine Validity -> Output CSV file|
|-Input File: 11658099 PNs|
|-Performance: 5941.72s @ 5000|

|06/02/2015|
|---------------------------|
|-First upload to GitHub|
|<b>-Currently working functions:</b>|
|  -Headers|
|  -Vlookup|
|  -ImportMulCSV|
|  -Population|
|  -Validator|
|  -Text2Column|
|<b>-Currently underdevelop functions:</b>|
|  -DataChecking|

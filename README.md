# Excel Instance Closer

## Overview
This VBA script continuously checks for running instances of Excel and closes them. It's useful for ensuring that no Excel applications remain open unnecessarily, which can free up system resources.

## Script Breakdown
1. **Infinite Loop:**
   Continuously runs to check for Excel instances.
   ```vbscript
   Do While True
   ```

2. **Variable Declaration**: Declares a variable to hold the Excel application object.
   ```vbscript
   Dim objExcel
   ```

3. **Error Handling**: Enables error handling to manage potential errors in the script.
   ```vbscript
   On Error Resume Next
   ```

4. **Check for Running Excel Instance**: Attempts to get an existing Excel application object. If none is found, it exits the loop.
   ```vbscript
   Set objExcel = GetObject(,"Excel.Application")
   If Err.Number <> 0 Then
    Exit Do
   End If
   ```

5. **Disable Alerts and Quit Excel**: Disables display alerts and quits the Excel application.
   ```vbscript
   On Error GoTo 0
   objExcel.DisplayAlerts = False
   objExcel.Quit
   Set objExcel = Nothing
   ```

5. **Loop Continuation**: Repeats the process indefinitely.
   ```vbscript
   Loop
   ```


## Usage

1. **Run the Script**: Execute the script in an appropriate VBA environment.



## License
This script is licensed under the MIT License. See the LICENSE file for details.



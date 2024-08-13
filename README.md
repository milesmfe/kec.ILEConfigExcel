# ILE Config Excel

<h2 id="1-excel-structure">Excel Structure</h2>

### Tabs (Sheets)

Both ILE and VE Excel Config Processors are divided into tabs, each of a specific type as identified by its highlighted colour (see table).

| **Type** | **Colour** | Function                                                                                               |
| -------------- | ---------------- | ------------------------------------------------------------------------------------------------------ |
| Input          | Green            | User is expected to paste data into sheet.                                                             |
| Output         | Blue             | Generated data is exported to sheet                                                                    |
| Processing     | Red              | Data is parsed through sheet, either containing formulae or blank and used for temporary data storage. |
| Mapping        | Grey             | Sheet contains a table mapping sets of values together, used in lookups.                              |

The process logic is very similar between ILE and VE however VE contains two additional tabs; Exchange Lookup and ILE Entry No. Map.

A full list of tabs for both ILE and VE is provided below.

#### ILE

| Name          | Type       | Description                                                                                                 |
| ------------- | ---------- | ----------------------------------------------------------------------------------------------------------- |
| INPUT         | Input      | Exported csv NAV data is pasted into column A of this sheet.                                                |
| OUTPUT        | Output     | Processed data is found on this sheet once processing is complete.                                          |
| ENTRY NO. MAP | Output     | Contains a list of every generated entry number mapped to its original (debug).                             |
| LINES         | Processing | *Formulae.* Calculates each line against the output header according to processing rules.                |
| BUDDY         | Processing | *Formulae.* Calculates each adjustment line against the output header according to processing rules.   |
| PROCESS       | Processing | *Formulae & Parse.* Calculates each line against the input header, adding temporary processing fields. |
| T2C           | Processing | *Parse.* Output location for the text-to-columns function.                                               |
| CUSTOMER MAP  | Mapping    | Maps customer IDs between two systems.                                                                      |

#### VE

| Name                        | Type              | Description                                                                                                                                                           |
| --------------------------- | ----------------- | --------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| INPUT                       | Input             | Exported csv NAV data is pasted into column A of this sheet.                                                                                                          |
| **ILE ENTRY NO. MAP** | **Input**   | **Generated table found in the Entry No. Map tab, after processing the corresponding ILE data, should be pasted under the header of the table in this sheet.** |
| OUTPUT                      | Output            | Processed data is found on this sheet once processing is complete.                                                                                                    |
| ENTRY NO. MAP               | Output            | Contains a list of every generated entry number mapped to its original (debug).                                                                                       |
| LINES                       | Processing        | *Formulae.* Calculates each line against the output header according to processing rules.                                                                          |
| BUDDY                       | Processing        | *Formulae.* Calculates each adjustment line against the output header according to processing rules.                                                             |
| PROCESS                     | Processing        | *Formulae & Parse.* Calculates each line against the input header, adding temporary processing fields.                                                           |
| T2C                         | Processing        | *Parse.* Output location for the text-to-columns function.                                                                                                         |
| **EXCHANGE LOOKUP**   | **Mapping** | **Maps an exchange rate to each day within the range of the input data.**                                                                                       |
| CUSTOMER MAP                | Mapping           | Maps customer IDs between two systems.                                                                                                                                |

### Parse (Blank)

There is only one blank processing tab, T2C, which is required to convert the csv input data, pasted into column A on the input tab, into a more useful format. Each line of the input data is formatted as follows:

`field1^field2^field3^field4^...fieldn`

where `n` is the total number of fields. Each line represents a record and each field is seperated by the delimiter character, `^`.

***Note:** **the total number of delimiter characters is equal to `n-1`.***

The first stage of processing uses the T2C sheet as an output to store the raw input data in rows and columns without a header, to later be used to populate the PROCESS tab with data.

The PROCESS tab is only labled as *Parse* due to the programatic transfer of data from T2C to PROCESS.

### Formulae (Table)

There are three processing tables each labeled according to their encapsulating tab's name. Each cell in tables LINES and BUDDY contains a formula pointing to the PROCESS table. These formulae implement the data rules as defined `<a href="#3-data-rules">`[here]`</a>`.

<h2 id="2-processing-logic">Processing Logic</h2>

Excel ILE config uses macros to move data between each tab, sometimes running functions or applying formulae. The macro triggers when a cell in column A is changed, then performs several tasks such as user input collection, data manipulation across different sheets, and updating tables with new values. The code consists of one main event handler and several subroutines and functions.

### ILE

* **Main handler:** `Worksheet_Change`

  ```vbnet
  Private Sub Worksheet_Change(ByVal Target As Range)
  ```
* **Checking change target**

  ```vbnet
  If Not Intersect(Target, Me.Columns("A")) Is Nothing Then
  ```

This line checks if the changed cell is in column A. If not, the macro exits.

* **Disabling Screen Updating and Calculation**
  ```vbnet
  Application.ScreenUpdating = False
  Application.Calculation = xlCalculationManual
  ```

This improves performance by preventing the screen from updating during the macro execution and disabling automatic recalculation of formulas.

* **Collecting User Inputs**

  ```vbnet
  firstEntryNo = GetUserInput("Enter the next available ILE Entry Number", "ILE  Entry No.")
  firstGlobalDimCode = GetUserInput("Enter the Global Dimension 1 Code", "Global Dimension 1 Code")
  currencyCode = GetUserInput("Enter the currency code: ", "Currency Code")
  ```

The macro prompts the user for three inputs: the next available ILE Entry Number, Global Dimension 1 Code, and currency code.

* **Updating Cells in the ****PROCESS**** table**

  ```vbnet
  Sheets("PROCESS").Range("N2").Formula = "=""" & firstGlobalDimCode & """"
  Sheets("PROCESS").Range("BW2").Formula = "=""" & currencyCode & """"
  ```

The macro updates cells N2 and BW2 in the PROCESS table with the user-provided Global Dimension 1 Code and Currency Code, respectively.

* **Processing Data**

  ```vbnet
  ProcessData
  ```

The `ProcessData` subroutine is called to perform data processing and manipulation.

* **Extending the LINES and BUDDY tables**

  ```vbnet
  ExtendTableWithFormulas "LINES", "lines"
  ExtendTableWithFormulas "BUDDY", "buddy"
  ```

The `ExtendTableWithFormulas` subroutine is called twice to ensure the formulas in the LINES and BUDDY tables are extended to match the data size.

* **Calculate Tables**

  ```vbnet
  Set processSheet = ThisWorkbook.Sheets("PROCESS")
  Set processTable = processSheet.ListObjects("process")
  Set linesSheet = ThisWorkbook.Sheets("LINES")
  Set linesTable = linesSheet.ListObjects("lines")
  Set buddySheet = ThisWorkbook.Sheets("BUDDY")
  Set buddyTable = buddySheet.ListObjects("buddy")

  processTable.Range.Calculate
  linesTable.Range.Calculate
  buddyTable.Range.Calculate
  ```

References to the process, lines, and buddy sheets/tables are set. The `Calculate` method is called on the ranges of these tables to refresh any calculations/formulas.

* **Generate Ouput Data**

  ```vbnet
  GenerateOutputData
  ```

The `GenerateOutputData` subroutine is called to produce the final output data by combining data from the “LINES” and “BUDDY” tables.

* **Updating Entry Numbers**

  ```vbnet
  UpdateOutputTableAndStore firstEntryNo
  ```

This subroutine updates entry numbers in the **OUTPUT** table and stores a mapping of these numbers in the Entry No. Map tab.

* **Informing the User**

  ```vbnet
  MsgBox "Processing Complete", vbInformation
  ```

A message box notifies the user that the processing is complete.

* **Re-enabling Screen Updating and Calculation**

  ```
  Application.ScreenUpdating = True
  Application.Calculation = xlCalculationAutomatic
  ```

Screen updating and automatic calculations are re-enabled.

### VE

* **Main handler:** `Worksheet_Change`

  ```vbnet
  Private Sub Worksheet_Change(ByVal Target As Range)
  ```
* **Checking change target**

  ```vbnet
  If Not Intersect(Target, Me.Columns("A")) Is Nothing Then
  ```

This line checks if the changed cell is in column A. If not, the macro exits.

* **Disabling Screen Updating and Calculation**
  ```vbnet
  Application.ScreenUpdating = False
  Application.Calculation = xlCalculationManual
  ```

This improves performance by preventing the screen from updating during the macro execution and disabling automatic recalculation of formulas.

* **Collecting User Inputs**

  ```vbnet
  firstEntryNo = GetUserInput("Enter the next available VE Entry Number", "VE  Entry No.")
  firstGlobalDimCode = GetUserInput("Enter the Global Dimension 1 Code", "Global Dimension 1 Code")
  ```

The macro prompts the user for three inputs: the next available VE Entry Number and Global Dimension 1 Code.

* **Updating Cells in the ****PROCESS**** table**

  ```vbnet
  Sheets("PROCESS").Range("V2").Formula = "=""" & firstGlobalDimCode & """"
  ```

The macro updates cell V2 in the PROCESS table with the user-provided Global Dimension 1 Code.

* **Processing Data**

  ```vbnet
  ProcessData
  ```

The `ProcessData` subroutine is called to perform data processing and manipulation.

* **Extending the LINES and BUDDY tables**

  ```vbnet
  ExtendTableWithFormulas "LINES", "lines"
  ExtendTableWithFormulas "BUDDY", "buddy"
  ```

The `ExtendTableWithFormulas` subroutine is called twice to ensure the formulas in the LINES and BUDDY tables are extended to match the data size.

* **Calculate Tables**

  ```vbnet
  Set processSheet = ThisWorkbook.Sheets("PROCESS")
  Set processTable = processSheet.ListObjects("process")
  Set linesSheet = ThisWorkbook.Sheets("LINES")
  Set linesTable = linesSheet.ListObjects("lines")
  Set buddySheet = ThisWorkbook.Sheets("BUDDY")
  Set buddyTable = buddySheet.ListObjects("buddy")

  processTable.Range.Calculate
  linesTable.Range.Calculate
  buddyTable.Range.Calculate
  ```

References to the process, lines, and buddy sheets/tables are set. The `Calculate` method is called on the ranges of these tables to refresh any calculations/formulas.

* **Generate Ouput Data**

  ```vbnet
  GenerateOutputData
  ```

The `GenerateOutputData` subroutine is called to produce the final output data by combining data from the “LINES” and “BUDDY” tables.

* **Updating Entry Numbers**

  ```vbnet
  UpdateOutputTableAndStore firstEntryNo
  ```

This subroutine updates entry numbers in the **OUTPUT** table and stores a mapping of these numbers in the Entry No. Map tab.

* **Informing the User**

  ```vbnet
  MsgBox "Processing Complete", vbInformation
  ```

A message box notifies the user that the processing is complete.

* **Re-enabling Screen Updating and Calculation**

  ```
  Application.ScreenUpdating = True
  Application.Calculation = xlCalculationAutomatic
  ```

Screen updating and automatic calculations are re-enabled.

### Subroutines

#### ProcessData

This subroutine copies data from column A in the INPUT sheet to the T2C sheet and then applies a Text-to-Columns operation on the T2C sheet with a specific delimiter. Finally, it copies the processed data to the PROCESS sheet and converts data in date columns into a date format.

1. **Set up worksheet references:**

   ```vbnet
   Set wsInput = ThisWorkbook.Sheets("INPUT")
   Set wsT2C = ThisWorkbook.Sheets("T2C")
   Set wsProcess = ThisWorkbook.Sheets("PROCESS")
   ```
2. **Clear the T2C sheet**

   ```vbnet
   wsT2C.Cells.Clear
   ```
3. **Copy data from INPUT to T2C**

   ```vbnet
   arrData = wsInput.Range("A1:A" & wsInput.Cells(wsInput.Rows.Count, "A").End(xlUp).Row).Value
   wsT2C.Range("A1").Resize(UBound(arrData, 1), 1).Value = arrData
   ```
4. **Apply Text-to-Columns**

   ```vbnet
   Set rng = wsT2C.Range("A1:A" & lastRow)
   rng.TextToColumns Destination:=wsT2C.Range("A1"), _
       DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, _
       ConsecutiveDelimiter:=False, Tab:=False, Semicolon:=False, Comma:=False, Space:=False, _
       Other:=True, OtherChar:="^"
   ```
5. **Copy processed data to PROCESS**

   ```vbnet
   wsProcess.Range("A2:BU" & lastRow + 1).Value = wsT2C.Range("A1:BU" & lastRow).Value
   ```
6. **Convert columns to dates**

   ```vbnet
   ConvertToDate wsProcess.Range("C2:C" & lastRow + 1)
   ConvertToDate wsProcess.Range("AF2:AF" & lastRow + 1)
   ConvertToDate wsProcess.Range("AR2:AR" & lastRow + 1)
   ```

#### ConvertToDate

Converts string representations of dates in a given range to proper date formats. Handles year values that are represented as two digits, assuming dates are in the 20th or 21st century.

1. **Load the range into an array**

   ```vbnet
   arrValues = rng.Value
   ```
2. **Process each value**

   ```vbnet
   For i = 1 To UBound(arrValues, 1)
       If IsDate(arrValues(i, 1)) Then
           ' Extract and reformat the date parts
           dateStr = CStr(arrValues(i, 1))
           dayPart = Left(dateStr, 2)
           monthPart = Mid(dateStr, 4, 2)
           yearPart = Right(dateStr, 2)
           If Val(yearPart) < 30 Then
               yearPart = "20" & yearPart
           Else
               yearPart = "19" & yearPart
           End If
           newDate = DateSerial(CInt(yearPart), CInt(monthPart), CInt(dayPart))
           arrValues(i, 1) = newDate
       End If
   Next i
   ```
3. **Write the array back to the range and apply format**

   ```vbnet
   rng.Value = arrValues
   rng.NumberFormat = "yyyy-mm-dd"
   ```

#### GetUserInput

Prompts the user for a non-empty string and returns it after validation.

1. **Prompt user and validate input**

   ```vbnet
   Do While Not isValidInput
       InputString = InputBox(prompt, title)
       If InputString <> "" Then
           isValidInput = True
       Else
           MsgBox "Invalid input. Please enter a non-empty string.", vbExclamation, "Invalid Input"
       End If
   Loop
   ```
2. **Return the validated string**

   ```vbnet
   GetUserInput = InputString
   ```

#### GenerateOutputData

Combines data from two tables (“LINES” and “BUDDY”) into a third table (“OUTPUT”) in an alternating pattern.

1. **Set worksheet and table references**

   ```vbnet
   Set ws1 = ThisWorkbook.Sheets("LINES")
   Set ws2 = ThisWorkbook.Sheets("BUDDY")
   Set ws3 = ThisWorkbook.Sheets("OUTPUT")
   Set table1 = ws1.ListObjects("lines")
   Set table2 = ws2.ListObjects("buddy")
   Set table3 = ws3.ListObjects("output")
   ```
2. **Initialise variables and output array**

   ```vbnet
   rowCount1 = table1.ListRows.Count
   rowCount2 = table2.ListRows.Count
   ReDim arrOutput(1 To rowCount1 + rowCount2, 1 To maxColumns)
   ```
3. **Loop through rows to combine data**

   ```vbnet
   For i = 1 To Application.WorksheetFunction.Max(rowCount1, rowCount2)
       If i <= rowCount1 Then
           For j = 1 To table1.ListColumns.Count
               arrOutput(outputRow, j) = table1.DataBodyRange(i, j).Value
           Next j
           outputRow = outputRow + 1
       End If
     
       If i <= rowCount2 Then
           For j = 1 To table2.ListColumns.Count
               arrOutput(outputRow, j) = table2.DataBodyRange(i, j).Value
           Next j
           outputRow = outputRow + 1
       End If
   Next i
   ```
4. **Write combined data to the OUTPUT table**

   ```vbnet
   If outputRow > 1 Then
       table3.Resize table3.Range.Resize(outputRow - 1, maxColumns)
       table3.DataBodyRange.Value = arrOutput
   End If
   ```

#### ExtendTableWithFormulas

Extends a target table by adding rows and applying formulas from the first row to all added rows.

1. **Set source and target references**

   ```vbnet
   Set sourceSheet = ThisWorkbook.Sheets("PROCESS")
   Set sourceTable = sourceSheet.ListObjects("process")
   Set targetSheet = ThisWorkbook.Sheets(targetSheetName)
   Set targetTable = targetSheet.ListObjects(targetTableName)
   ```
2. **Calculate rows to add and resize table**

   ```vbnet
   rowsToAdd = sourceTable.ListRows.Count - targetTable.ListRows.Count
   If rowsToAdd > 0 Then
       targetTable.Resize targetTable.Range.Resize(targetTable.ListRows.Count + rowsToAdd)
   End If
   ```
3. **Apply formulas to the new rows**

   ```vbnet
   If targetTable.ListRows.Count > 1 Then
       lastDataRow = targetTable.Range.Rows.Count - 1
       For i = 1 To targetTable.ListColumns.Count
           formulaCell = targetTable.DataBodyRange(1, i)
           If formulaCell.HasFormula Then
               targetTable.DataBodyRange(2, i).Resize(rowsToAdd).Formula = formulaCell.Formula
           End If
       Next i
   End If
   ```

#### UpdateOutputTableAndStore

Updates the OUTPUT table by replacing the first column values with incrementing entry numbers starting from a specified number. It also stores the mapping between the original data and the new entry numbers in the ENTRY NO. MAP table.

1. **Declare variables**

   ```vbnet
       Dim wsOutput As Worksheet
       Dim wsEntryNoMap As Worksheet
       Dim outputTable As ListObject
       Dim entryNoMapTable As ListObject
       Dim firstEntryNoValue As Long
       Dim rowCount As Long
       Dim i As Long
       Dim outputArray As Variant
       Dim entryNoMapArray As Variant
   ```
2. **Set worksheet and table references**

   ```vbnet
   Set wsOutput = ThisWorkbook.Sheets("OUTPUT")
   Set outputTable = wsOutput.ListObjects("output")
   Set wsEntryNoMap = ThisWorkbook.Sheets("ENTRY NO. MAP")
   Set entryNoMapTable = wsEntryNoMap.ListObjects("entryNoMap")
   ```
3. **Assign and validate firstEntryNo**

   ```vbnet
   firstEntryNoValue = CLng(firstEntryNo)
   ```
4. **Count rows in the output table**

   ```vbnet
       rowCount = outputTable.ListRows.Count
       If rowCount = 0 Then Exit Sub
   ```
5. **Read and initialise arrays**

   ```vbnet
   outputArray = outputTable.ListColumns(1).DataBodyRange.Value
   ReDim entryNoMapArray(1 To rowCount, 1 To 2)
   ```
6. **Populate arrays with data**

   ```vbnet
   For i = 1 To rowCount
       entryNoMapArray(i, 1) = outputArray(i, 1)
       entryNoMapArray(i, 2) = CStr(firstEntryNoValue + i - 1)
       outputArray(i, 1) = entryNoMapArray(i, 2)
   Next i
   ```
7. **Write data back to tables**

   ```vbnet
   outputTable.ListColumns(1).DataBodyRange.Value = outputArray
   entryNoMapTable.DataBodyRange.ClearContents
   entryNoMapTable.DataBodyRange.Resize(rowCount, 2).Value = entryNoMapArray
   ```

<h2 id="3-data-rules">Data Rules</h2>

### ILE


### VE

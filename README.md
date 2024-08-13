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

There are three processing tables each labeled according to their encapsulating tab's name. Each cell in tables LINES and BUDDY contains a formula pointing to the PROCESS table. These formulae implement the data rules as defined <a href="#3-data-rules">[here]</a>.

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

Lorem Ipsum.

<h2 id="3-data-rules">Data Rules</h2>

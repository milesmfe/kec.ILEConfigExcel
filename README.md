# ILE Config Excel

## Excel Structure

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

There are three processing tables each labeled according to their encapsulating tab's name. Each cell in tables LINES and BUDDY contains a formula pointing to the PROCESS table. These formulae implement the processing logic as defined below.

## Processing Logic

Lorem ipsum.

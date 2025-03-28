# XML Generator for Saxo Trading Data

## Overview

This Python script processes trading data exported from the Saxo platform (in `.xlsx` format) and converts it into an XML file suitable for tax reporting. It also incorporates exchange rate data from an XML source to convert values into EUR.

## Features

- Parses an Excel file containing Saxo trading data.
- Reads exchange rates from an XML file.
- Converts transaction values to EUR using the nearest historical exchange rate.
- Generates a structured XML file for tax reporting.

## Requirements

Before running the script, ensure you have the following installed:

- Python 3.x
- Required Python libraries:
  ```sh
  pip install pandas openpyxl
  ```

## Usage

1. Prepare the necessary files:

   - Trading data exported from Saxo as `taxexample.xlsx`.
   - Exchange rate data as `usdtoeur.xml`.

2. Run the script:

   ```sh
   python script.py
   ```

3. Enter the required taxpayer and KDVP details when prompted.

4. The generated XML file (`output.xml`) will be created in the script's directory.

## Input Data Format

### **Excel File (`taxexample.xlsx`):**

The script expects the trading data in an Excel file with the following columns:

- **Date of Sale**
- **Date of Purchase**
- **Stock Name / Code**
- **Quantity**
- **Purchase Price (in original currency)**
- **Sale Price (in original currency)**

### **Exchange Rate XML (`usdtoeur.xml`):**

- The XML file should contain exchange rate data structured according to the European Central Bank's (ECB) format.
- The script extracts the nearest historical exchange rate for currency conversion.

## Output

- The script generates an XML file (`output.xml`) containing structured tax information.

## Example Output

```xml
<Envelope xmlns="http://edavki.durs.si/Documents/Schemas/Doh_KDVP_9.xsd" xmlns:edp="http://edavki.durs.si/Documents/Schemas/EDP-Common-1.xsd">
    <edp:Header>
        <edp:taxpayer>
            <edp:taxNumber>12345678</edp:taxNumber>
            <edp:name>John Doe</edp:name>
            ...
        </edp:taxpayer>
    </edp:Header>
    <body>
        <Doh_KDVP>
            <KDVP>
                <SecurityCount>5</SecurityCount>
                <KDVPItem>
                    <ItemID>1</ItemID>
                    <Name>Stock A</Name>
                    ...
                </KDVPItem>
            </KDVP>
        </Doh_KDVP>
    </body>
</Envelope>
```

## Customization

- Modify the `generate_xml_from_excel` function to adjust the XML output format.
- Update the `parse_exchange_rates` function if using a different exchange rate XML structure.

## Author

This script was created for automating tax reporting for Saxo trading exports. Contributions and improvements are welcome!

# Excel to JSON Bank Statement Parser

A dynamic Java-based utility to parse unstructured bank statement Excel files (`.xls` and `.xlsx`) and convert them into structured, easy-to-read JSON files.

## Features

- **Dynamic Table Detection**: Automatically scans entire Excel sheets to dynamically locate the transaction data header regardless of preceding rows or varying metadata (e.g., bank logo, summary, user details above the tables).
- **Format Agnostic**: Relies on specific keywords (Date, Transaction, Amount, Debit, Credit, Balance, etc.) to map out column boundaries.
- **Robust Read Capability**: Powered by Apache POI to support both modern `.xlsx` and legacy `.xls` statement formats seamlessly.
- **Smart Data Extraction**: Avoids empty gaps and skips unrelated footer sections. Converts date strings smoothly into JSON-native outputs using Jackson Databind.
- **Automatic Output**: Dumps the converted output directly onto the console and automatically writes to a JSON file alongside the original input file.

## Prerequisites

- **Java 11** or higher
- **Maven 3.x**

## Build Instructions 

Use Maven to build a single FAT jar containing all dependencies.

```bash
mvn clean package
```

This ensures the `maven-shade-plugin` compiles the executable `xlsx-to-json-parser-1.0-SNAPSHOT.jar` inside the `target/` directory.

## Usage

### 1. Using the Batch File (Windows) 

To make running simple, you can use the provided Windows batch script `run-parser.bat`. If the project JAR hasn't been built yet, the script will automatically invoke `mvn package` before executing.

```cmd
run-parser.bat <BankName> <ExcelFilePath>
```

**Example:**
```cmd
.\run-parser.bat "HDFC" "xlsx-files\Bank-Statement-Template-1-TemplateLab.xlsx"
```

### 2. Using Java Command Line

Run the generated FAT Jar directly with Java:

```bash
java -jar target/xlsx-to-json-parser-1.0-SNAPSHOT.jar <BankName> <ExcelFilePath>
```

**Example:**
```bash
java -jar target/xlsx-to-json-parser-1.0-SNAPSHOT.jar "ICICI" "xlsx-files/Bank-Statement-Template-2-TemplateLab.xlsx"
```

## Output

Upon successful execution, a structured JSON file will be generated in the same directory as the original `.xlsx` file. For example, if you ran `Statement.xlsx`, you'll receive a `Statement.json` file like this:

```json
{
  "bankName" : "ICICI",
  "count" : 3,
  "transactions" : [ {
    "date" : "mm/dd/yyyy",
    "payment type" : "Fast Payment",
    "detail" : "Amazon",
    "paid out" : 132.3,
    "balance" : 8180.99
  }, {
    "date" : "mm/dd/yyyy",
    "payment type" : "BACS",
    "detail" : "Toyota Online",
    "paid out" : 10525.4,
    "balance" : 12633.64
  }, {
    "date" : "Note:"
  } ]
}
```

## How It Works

1. **Workbook Loading**: Apache POI opens the given workbook file.
2. **Scoring the Sheet**: Evaluates each row by counting column hits from a predetermined set of keywords (e.g., *withdrawal, deposit, description, narration, amount*).
3. **Establishing Boundary**: Once a threshold matches and the row with the maximum hit-score is found, it maps the index of columns mathematically.
4. **Data Iteration**: Processes subsequent lines up untill blank patterns appear frequently, marking the transactions structure directly into mapped keys via Jackson serialization.

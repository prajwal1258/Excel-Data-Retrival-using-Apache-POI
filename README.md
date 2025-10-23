Java Excel Data Retriever using Apache POI

A simple Java project demonstrating how to read and retrieve specific data from an Excel (.xlsx) file using the Apache POI library.

This example provides a use case for parsing an Excel sheet to find and extract data from a specific row based on a key value (e.g., an Employee ID).

Features

File Handling: Opens an .xlsx workbook and accesses a specific sheet using FileInputStream and XSSFWorkbook.

Data Iteration: Iterates through all rows and cells in the sheet.

Safe Data Formatting: Uses DataFormatter to correctly read and convert cell values (e.g., numbers, strings) into a consistent string format.

Targeted Data Retrieval: Implements logic to:

Find and print the header row.

Search for a specific row where a cell matches a target ID (e.g., "E4015").

Print all subsequent cell data from that matching row.

Core Dependencies

Apache POI (Binary Artifacts)

poi-ooxml (for .xlsx file support)


# Spring Boot Application for Modifying Excel Files

This repository contains a **Spring Boot** application that demonstrates how to read, modify, and save data in Excel files using **Apache POI**. The application is particularly useful for handling large Excel files (easily over 9,000 rows) by automating cell checks and updates. 

## Table of Contents
1. [Overview](#overview)
2. [Features](#features)
3. [Technologies Used](#technologies-used)
4. [Prerequisites](#prerequisites)
5. [How to Run](#how-to-run)
6. [Code Explanation](#code-explanation)
7. [Example Use Case](#example-use-case)
8. [Troubleshooting](#troubleshooting)
9. [License](#license)

---

## Overview
This project demonstrates:
- Reading an Excel file from a given path.
- Checking if certain cells are empty.
- Copying values from one cell to another if needed.
- Saving the modified data to a new Excel file.

It uses **Spring Boot** to run the application via the `CommandLineRunner` interface, enabling the program to execute the Excel processing logic immediately upon startup.

---

## Features
- **Automatic Cell Checking**: Quickly checks if specified cells (e.g., city, address) are empty.
- **Cell Copy Logic**: Copies the non-empty cell content to the empty cell if one of them is missing.
- **Batch Processing**: Efficiently processes thousands of rows without manual intervention.
- **Logging**: Prints a summary of every modified row and the total changes made.

---

## Technologies Used
- **Java 8+** (or any higher LTS version)
- **Spring Boot** (for easy setup and CLI execution)
- **Apache POI** (for Excel file processing)
- **Maven** (as a build automation tool)

---

## Prerequisites
- **Java 8+** installed on your machine.
- **Maven** installed (if you intend to build and run from the command line).
- An IDE like IntelliJ, Eclipse, or Visual Studio Code (optional, but recommended).
- **Git** (optional, if you are cloning this repository).

---

## How to Run
1. **Clone the Repository** (or download the ZIP):
   ```bash
   git clone https://github.com/abdulrahmanOmar1/Spring-Boot-application-that-modifies-Excel-files.git
   ```
2. **Navigate to the Project Folder**:
   ```bash
   cd Spring-Boot-application-that-modifies-Excel-files
   ```
3. **Build the Project** (using Maven):
   ```bash
   mvn clean install
   ```
4. **Run the Application**:
   ```bash
   mvn spring-boot:run
   ```
   Alternatively, you can run the generated `.jar` file located in the `target` folder:
   ```bash
   java -jar target/your-jar-file-name.jar
   ```

When the application starts, it will execute the Excel processing logic in the `run` method of the `CommandLineRunner`.

---

## Code Explanation
Below is a brief explanation of the core logic found in **ExaelApplication** (or **ExcelProcessor**, depending on which class you use):

1. **Input/Output File Paths**:
   ```java
   String inputFilePath = "C:/path/to/your/input.xlsx";
   String outputFilePath = "C:/path/to/your/output.xlsx";
   ```
   - `inputFilePath`: The existing Excel file to read.
   - `outputFilePath`: Where the processed Excel file will be saved.

2. **Reading the Workbook**:
   ```java
   try (FileInputStream fis = new FileInputStream(new File(inputFilePath));
        Workbook workbook = new XSSFWorkbook(fis)) {
       Sheet sheet = workbook.getSheetAt(0);
       ...
   }
   ```
   - **FileInputStream** loads the Excel file into memory.
   - **XSSFWorkbook** is used for `.xlsx` format.  
   - `sheet` is the first sheet in the workbook.

3. **Iterating Over Rows**:
   ```java
   for (Row row : sheet) {
       // Access specific cells, e.g. row.getCell(1), row.getCell(5)
       ...
   }
   ```
   - Loops through each row in the sheet.

4. **Logic for Empty Cells**:
   ```java
   if (isCellEmpty(cityCell)) {
       // If city is empty, copy from address or default to "رام الله"
   }
   if (isCellEmpty(addressCell)) {
       // If address is empty, copy from city
   }
   ```
   - Checks if a particular cell is empty using a helper method `isCellEmpty`.

5. **Saving the Modified Data**:
   ```java
   try (FileOutputStream fos = new FileOutputStream(new File(outputFilePath))) {
       workbook.write(fos);
   }
   ```
   - Writes the changes to a new Excel file at the specified output path.

6. **Logging & Counting**:
   ```java
   List<String> modifiedRecords = new ArrayList<>();
   int modifiedCount = 0;

   // Each time a cell is updated, add a log message and increment modifiedCount
   ```
   - Keeps track of every modification made for easy review in the console.

---

## Example Use Case
Suppose you have an Excel file with thousands of records where:
- The **City** column is at index `1`.
- The **Address** column is at index `5`.

You need to ensure:
- If **City** is empty but **Address** is filled, copy **Address** into **City**.
- If **Address** is empty but **City** is filled, copy **City** into **Address**.
- Or set default values (e.g., "Ramallah") for both if both are empty.

Simply specify the file paths in the code and run the application. The program will:
1. Open the existing Excel file.
2. Check every row, updating cells according to the logic.
3. Print out each modified row in the console.
4. Save the final updated workbook to the output file.

---

## Troubleshooting
- **File Not Found**: Ensure the `inputFilePath` is correct and points to an existing `.xlsx` file.
- **Permission Issues**: Make sure you have permission to read from the input file and write to the output file directory.
- **Corrupted File**: If you get an error about a corrupted file, verify the file is a valid `.xlsx` file.


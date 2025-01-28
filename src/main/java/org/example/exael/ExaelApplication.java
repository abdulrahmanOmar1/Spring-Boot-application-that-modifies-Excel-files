package org.example.exael;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.boot.CommandLineRunner;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

@SpringBootApplication
public class ExaelApplication implements CommandLineRunner {

    public static void main(String[] args) {
        SpringApplication.run(ExaelApplication.class, args);
    }

    @Override
    public void run(String... args) throws Exception {
        String inputFilePath = "C:/Users/abd.almahmoud/Downloads/Test.xlsx";
        String outputFilePath = "C:/Users/abd.almahmoud/Downloads/output2.xlsx";

        List<String> modifiedRecords = new ArrayList<>();
        int modifiedCount = 0;

        try (FileInputStream fis = new FileInputStream(new File(inputFilePath));
             Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = workbook.getSheetAt(0);

            for (Row row : sheet) {
                Cell cityCell = row.getCell(1);
                Cell addressCell = row.getCell(5);

                if (isCellEmpty(cityCell)) {
                    if (!isCellEmpty(addressCell)) {
                        cityCell = row.createCell(1);
                        cityCell.setCellValue(getCellValueAsString(addressCell));
                        modifiedRecords.add("Modified row " + row.getRowNum() + ": Set city to " + getCellValueAsString(addressCell));
                        modifiedCount++;
                    }
                } else if (isCellEmpty(addressCell)) {
                    if (!isCellEmpty(cityCell)) {
                        addressCell = row.createCell(5);
                        addressCell.setCellValue(getCellValueAsString(cityCell));
                        modifiedRecords.add("Modified row " + row.getRowNum() + ": Set address to " + getCellValueAsString(cityCell));
                        modifiedCount++;
                    }
                }
            }

            try (FileOutputStream fos = new FileOutputStream(new File(outputFilePath))) {
                workbook.write(fos);
            }

            System.out.println("File processed and saved successfully!");

            for (String record : modifiedRecords) {
                System.out.println(record);
            }

            System.out.println("Total modified records: " + modifiedCount);

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private boolean isCellEmpty(Cell cell) {
        return cell == null || cell.getCellType() == CellType.BLANK;
    }

    private String getCellValueAsString(Cell cell) {
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                return String.valueOf(cell.getNumericCellValue());
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                return cell.getCellFormula();
            default:
                return "";
        }
    }
}
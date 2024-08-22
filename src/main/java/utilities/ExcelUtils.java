package utilities;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class ExcelUtils {

    /**
     * Filter the Excel sheet by the truck name and write the filtered data to a new Excel file.
     * The output file will contain only the rows where the "Truck" column matches the given truck name.
     * The output file will have the same header row as the input file.
     * If the "Truck" column is not found, it will print a message and return without writing the output file.
     * If the truck name is not found in any row, the output file will be created but it will be empty.
     * If the input file does not exist or is not a valid Excel file, it will throw an exception.
     * If the output file already exists, it will be overwritten.
     *
     * @param inputFilePath The path to the input Excel file.
     * @param sheetNameToFilter The name of the sheet to filter.
     * @param outputFilePath The path to the output Excel file.
     * @param truckFullName The full name of the truck to filter by.
     */
    public void filterExcelByTruck(String inputFilePath,String sheetNameToFilter, String outputFilePath, String truckFullName) {

        try (FileInputStream fileInputStream = new FileInputStream(inputFilePath);
             XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream)) {

            // Get the first sheet
            XSSFSheet sheet = workbook.getSheet(sheetNameToFilter);

            // Identify the "Truck" column
            Row headerRow = sheet.getRow(0);
            int truckColumnIndex = -1; // Initialize to -1 to indicate not found
            for (Cell cell : headerRow) {
                //if the cell value is equal to "Truck" then set the truckColumnIndex to the index of the cell
                if (cell.getStringCellValue().equalsIgnoreCase("Truck")) {
                    truckColumnIndex = cell.getColumnIndex();
                    break;
                }
            }
            //if the truckColumnIndex is still -1 then print a message and return from the method as the column is not found
            if (truckColumnIndex == -1) {
                System.out.println("Truck column not found.");
                return;
            }

            // Create a new workbook and sheet for output
            XSSFWorkbook outputWorkbook = new XSSFWorkbook();
            XSSFSheet outputSheet = outputWorkbook.createSheet("FilteredData");

            // Copy the header row to the output sheet
            Row outputHeaderRow = outputSheet.createRow(0);
            for (Cell cell : headerRow) {
                Cell newCell = outputHeaderRow.createCell(cell.getColumnIndex());
                newCell.setCellValue(cell.getStringCellValue());
            }

            // Loop through the rows and check the Truck column
            int outputRowIndex = 1; // Start from the second row (first row is the header)
            for (int i = 1; i <= sheet.getLastRowNum(); i++) { // Start from 1 to skip the header row
                Row row = sheet.getRow(i);
                if (row != null) {
                    Cell truckCell = row.getCell(truckColumnIndex);
                    if (truckCell != null && truckCell.getStringCellValue().equalsIgnoreCase(truckFullName)) {
                        // Copy the entire row to the output sheet
                        Row outputRow = outputSheet.createRow(outputRowIndex++);
                        for (Cell cell : row) {
                            Cell newCell = outputRow.createCell(cell.getColumnIndex());
                            copyCellValueAndStyle(cell, newCell, outputWorkbook);
                        }
                    }
                }
            }

            // Write the output to a new Excel file
            try (FileOutputStream fileOutputStream = new FileOutputStream(outputFilePath)) {
                outputWorkbook.write(fileOutputStream);
            }

            System.out.println("Filtered rows have been written to the output file.");

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private void copyCellValueAndStyle(Cell sourceCell, Cell targetCell, XSSFWorkbook outputWorkbook) {
        switch (sourceCell.getCellType()) {
            case STRING:
                targetCell.setCellValue(sourceCell.getStringCellValue());
                break;
            case NUMERIC:
                // Check if the cell contains a date value
                if (DateUtil.isCellDateFormatted(sourceCell)) {
                    targetCell.setCellValue(sourceCell.getDateCellValue());
                } else {
                    // Copy the numeric value
                    targetCell.setCellValue(sourceCell.getNumericCellValue());
                }
                break;
            case BOOLEAN:
                // Copy the boolean value
                targetCell.setCellValue(sourceCell.getBooleanCellValue());
                break;
            case FORMULA:
                // Copy the formula
                targetCell.setCellFormula(sourceCell.getCellFormula());
                break;
            default:
                break;
        }

        // Copy the cell style, including the number format
        CellStyle newCellStyle = outputWorkbook.createCellStyle();
        newCellStyle.cloneStyleFrom(sourceCell.getCellStyle());
        targetCell.setCellStyle(newCellStyle);
    }



    /**
     * Generate a report by calculating the sum of "Rate", "Gross Pay", and "Total" columns.
     * The output will be written to the same Excel file.
     * Note: This method assumes that the input Excel file already contains the "FilteredData" sheet.
     * If the sheet does not exist, it will throw an exception.
     * If the required columns are not found, it will print a message and return.
     * If the required columns are found, it will calculate the sums and add them to the last row.
     * Finally, it will write the output back to the same Excel file.
     * @param filePath The path to the Excel file.
     */
    public void generateReport(String filePath) {
        try (FileInputStream fileInputStream = new FileInputStream(filePath);
             XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream)) {

            // Get the first sheet
            XSSFSheet sheet = workbook.getSheet("FilteredData");

            // Identify the column indices for "Rate", "Gross Pay", and "Total"
            Row headerRow = sheet.getRow(0);
            int rateColumnIndex = -1;
            int grossPayColumnIndex = -1;
            int totalColumnIndex = -1;
            for (Cell cell : headerRow) {
                String header = cell.getStringCellValue().trim();
                if (header.equalsIgnoreCase("Rate")) {
                    rateColumnIndex = cell.getColumnIndex();
                } else if (header.equalsIgnoreCase("Gross Pay")) {
                    grossPayColumnIndex = cell.getColumnIndex();
                } else if (header.equalsIgnoreCase("Total")) {
                    totalColumnIndex = cell.getColumnIndex();
                }
            }

            if (rateColumnIndex == -1 || grossPayColumnIndex == -1 || totalColumnIndex == -1) {
                System.out.println("Required column(s) not found.");
                return;
            }

            // Calculate sums for the "Rate", "Gross Pay", and "Total" columns
            double rateSum = 0;
            double grossPaySum = 0;
            double totalSum = 0;

            CellStyle rateCellStyle = null;
            CellStyle grossPayCellStyle = null;
            CellStyle totalCellStyle = null;

            for (int i = 1; i <= sheet.getLastRowNum(); i++) { // Start from 1 to skip the header row
                Row row = sheet.getRow(i);
                if (row != null) {
                    rateSum += getNumericCellValue(row, rateColumnIndex);
                    grossPaySum += getNumericCellValue(row, grossPayColumnIndex);
                    totalSum += getNumericCellValue(row, totalColumnIndex);

                    // Store the first cell styles to use for the sum row
                    //rateCellStyle is initially set to null when the variable is declared.
                    //The if statement checks if rateCellStyle is still null.
                    //The reason for this check is to ensure that the cell style is only
                    // retrieved and stored once, specifically from the first non-header row that has a "Rate" value.
                    if (rateCellStyle == null) {
                        rateCellStyle = row.getCell(rateColumnIndex).getCellStyle();
                    }
                    if (grossPayCellStyle == null) {
                        grossPayCellStyle = row.getCell(grossPayColumnIndex).getCellStyle();
                    }
                    if (totalCellStyle == null) {
                        totalCellStyle = row.getCell(totalColumnIndex).getCellStyle();
                    }
                }
            }

            // Create a new row for the sums
            Row sumRow = sheet.createRow(sheet.getLastRowNum() + 1);

            // Set the summed values and apply the copied cell styles
            Cell rateSumCell = sumRow.createCell(rateColumnIndex);
            rateSumCell.setCellValue(rateSum);
            if (rateCellStyle != null) {
                rateSumCell.setCellStyle(rateCellStyle);//set the cell style of the rateSumCell to the rateCellStyle
            }

            Cell grossPaySumCell = sumRow.createCell(grossPayColumnIndex);
            grossPaySumCell.setCellValue(grossPaySum);
            if (grossPayCellStyle != null) {
                grossPaySumCell.setCellStyle(grossPayCellStyle);//set the cell style of the grossPaySumCell to the grossPayCellStyle
            }

            Cell totalSumCell = sumRow.createCell(totalColumnIndex);
            totalSumCell.setCellValue(totalSum);
            if (totalCellStyle != null) {
                totalSumCell.setCellStyle(totalCellStyle);//set the cell style of the totalSumCell to the totalCellStyle
            }

            // Write the output back to the same Excel file
            try (FileOutputStream fileOutputStream = new FileOutputStream(filePath)) {
                workbook.write(fileOutputStream);
            }

            System.out.println("Sum of 'Rate', 'Gross Pay', and 'Total' columns has been added to the last row.");

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static double getNumericCellValue(Row row, int columnIndex) {
        Cell cell = row.getCell(columnIndex);
        if (cell != null && cell.getCellType() == CellType.NUMERIC) {
            return cell.getNumericCellValue();
        }
        return 0;
    }

}

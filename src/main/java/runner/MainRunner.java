package runner;

import utilities.ExcelUtils;

public class MainRunner {
    public static void main(String[] args) {
        ExcelUtils excelUtils = new ExcelUtils();

        try {
            // Read truck names from the Excel file
            String[] trucks = excelUtils.readTruckNamesFromExcel("newfile.xlsx", "Sheet1", "Truck");

            // Filter by each truck and create a report
            excelUtils.filterExcelByTruck("newfile.xlsx", "Sheet1", "output.xlsx", "Truck");
            excelUtils.generateReport("output.xlsx", trucks );
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
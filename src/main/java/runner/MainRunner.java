package runner;

import utilities.ExcelUtils;

public class MainRunner {
    public static void main(String[] args) {
        ExcelUtils excelUtils = new ExcelUtils();

        try {
            excelUtils.filterExcelByTruck("example of report.xlsx","Sheet1", "output.xlsx", "Truck #001");
            excelUtils.generateReport("output.xlsx");
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}

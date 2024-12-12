import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;

public class ExcelUtil {

    private Workbook workbook;
    private String filePath;

    public ExcelUtil(String filePath) throws IOException {
        this.filePath = filePath;
        try (FileInputStream fis = new FileInputStream(filePath)) {
            if (filePath.endsWith(".xlsx")) {
                this.workbook = new XSSFWorkbook(fis);
            } else if (filePath.endsWith(".xls")) {
                this.workbook = new HSSFWorkbook(fis);
            } else {
                throw new IllegalArgumentException("The file format is not supported. Please provide an .xlsx or .xls file.");
            }
        }
    }

    /**
     * Reads data from a specified sheet dynamically.
     * @param sheetName the name of the sheet to read.
     * @return a map where each key is a header (from column 1),
     *         and each value is a map of test case names (from row 1) to their respective values.
     */
    public Map<String, Map<String, String>> readSheetData(String sheetName) {
        Sheet sheet = workbook.getSheet(sheetName);
        if (sheet == null) {
            throw new IllegalArgumentException("Sheet with name " + sheetName + " does not exist.");
        }

        Map<String, Map<String, String>> dataMap = new HashMap<>();
        Iterator<Row> rowIterator = sheet.iterator();

        if (!rowIterator.hasNext()) {
            throw new IllegalStateException("The sheet is empty.");
        }

        // Read headers (column 1 values)
        Row headerRow = rowIterator.next();
        Iterator<Cell> cellIterator = headerRow.cellIterator();
        Map<Integer, String> testCaseMap = new HashMap<>();

        while (cellIterator.hasNext()) {
            Cell cell = cellIterator.next();
            if (cell.getColumnIndex() > 0) { // Skip first column (headers)
                testCaseMap.put(cell.getColumnIndex(), getCellValue(cell));
            }
        }

        // Read data dynamically
        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();
            Cell headerCell = row.getCell(0); // First column is the header
            if (headerCell == null) continue;

            String header = getCellValue(headerCell);
            if (!dataMap.containsKey(header)) {
                dataMap.put(header, new HashMap<>());
            }

            Map<String, String> testCaseValues = dataMap.get(header);
            for (int colIndex : testCaseMap.keySet()) {
                Cell cell = row.getCell(colIndex);
                String testCase = testCaseMap.get(colIndex);
                testCaseValues.put(testCase, cell != null ? getCellValue(cell) : "");
            }
        }

        return dataMap;
    }

    /**
     * Reads data from all sheets dynamically.
     * @return a map where each key is the sheet name, and each value is the sheet data map.
     */
    public Map<String, Map<String, Map<String, String>>> readAllSheetsData() {
        Map<String, Map<String, Map<String, String>>> allData = new HashMap<>();
        for (Sheet sheet : workbook) {
            allData.put(sheet.getSheetName(), readSheetData(sheet.getSheetName()));
        }
        return allData;
    }

    /**
     * Gets the row count of a sheet.
     * @param sheetName the name of the sheet.
     * @return the number of rows in the sheet.
     */
    public int getRowCount(String sheetName) {
        Sheet sheet = workbook.getSheet(sheetName);
        if (sheet == null) {
            throw new IllegalArgumentException("Sheet with name " + sheetName + " does not exist.");
        }
        return sheet.getPhysicalNumberOfRows();
    }

    /**
     * Gets the column count of a sheet.
     * @param sheetName the name of the sheet.
     * @return the number of columns in the first row of the sheet.
     */
    public int getColCount(String sheetName) {
        Sheet sheet = workbook.getSheet(sheetName);
        if (sheet == null) {
            throw new IllegalArgumentException("Sheet with name " + sheetName + " does not exist.");
        }
        Row firstRow = sheet.getRow(0);
        if (firstRow == null) {
            return 0;
        }
        return firstRow.getPhysicalNumberOfCells();
    }

    /**
     * Updates a cell value based on header name and sheet name.
     * @param sheetName the name of the sheet.
     * @param headerName the header name in the first column.
     * @param testCaseName the test case name in the first row.
     * @param newValue the new value to set.
     */
    public void updateExcel(String sheetName, String headerName, String testCaseName, String newValue) {
        Sheet sheet = workbook.getSheet(sheetName);
        if (sheet == null) {
            throw new IllegalArgumentException("Sheet with name " + sheetName + " does not exist.");
        }

        int headerRowIndex = -1;
        int testCaseColIndex = -1;

        // Find the header row index
        Iterator<Row> rowIterator = sheet.iterator();
        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();
            Cell cell = row.getCell(0);
            if (cell != null && getCellValue(cell).equalsIgnoreCase(headerName)) {
                headerRowIndex = row.getRowNum();
                break;
            }
        }

        if (headerRowIndex == -1) {
            throw new IllegalArgumentException("Header with name " + headerName + " does not exist.");
        }

        // Find the test case column index
        Row headerRow = sheet.getRow(0);
        for (Cell cell : headerRow) {
            if (getCellValue(cell).equalsIgnoreCase(testCaseName)) {
                testCaseColIndex = cell.getColumnIndex();
                break;
            }
        }

        if (testCaseColIndex == -1) {
            throw new IllegalArgumentException("Test case with name " + testCaseName + " does not exist.");
        }

        // Update the cell value
        Row rowToUpdate = sheet.getRow(headerRowIndex);
        Cell cellToUpdate = rowToUpdate.getCell(testCaseColIndex);
        if (cellToUpdate == null) {
            cellToUpdate = rowToUpdate.createCell(testCaseColIndex);
        }
        cellToUpdate.setCellValue(newValue);

        try (FileOutputStream fos = new FileOutputStream(filePath)) {
            workbook.write(fos);
        } catch (IOException e) {
            e.printStackTrace();
            throw new RuntimeException("Failed to write to the Excel file.");
        }
    }

    /**
     * Utility method to get cell value as a string.
     */
    private String getCellValue(Cell cell) {
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return cell.getDateCellValue().toString();
                } else {
                    return String.valueOf(cell.getNumericCellValue());
                }
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                return cell.getCellFormula();
            case BLANK:
                return "";
            default:
                return "";
        }
    }

    /**
     * Closes the workbook resource.
     */
    public void close() throws IOException {
        if (workbook != null) {
            workbook.close();
        }
    }

    public static void main(String[] args) {
        try {
            ExcelUtil excelUtil = new ExcelUtil("path/to/your/excel/file.xlsx");

            // Reading data from a specific sheet
            Map<String, Map<String, String>> sheetData = excelUtil.readSheetData("Sheet1");
            System.out.println("Sheet1 Data: " + sheetData);

            // Reading data from all sheets
            Map<String, Map<String, Map<String, String>>> allData = excelUtil.readAllSheetsData();
            System.out.println("All Sheets Data: " + allData);

            // Getting row and column counts
            System.out.println("Row Count: " + excelUtil.getRowCount("Sheet1"));
            System.out.println("Column Count: " + excelUtil.getColCount("Sheet1"));

            // Updating a cell value
            excelUtil.updateExcel("Sheet1", "Header1", "TestCase1", "NewValue");

            excelUtil.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}

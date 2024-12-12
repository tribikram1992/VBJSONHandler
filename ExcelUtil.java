import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
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

    public Map<String, Map<String, String>> readSheetData(String sheetName) {
        Sheet sheet = workbook.getSheet(sheetName);
        if (sheet == null) {
            throw new IllegalArgumentException("Sheet with name " + sheetName + " does not exist.");
        }

        Map<String, Map<String, String>> dataMap = new HashMap<>();
        Map<Integer, String> testCaseMap = new HashMap<>();

        Row firstRow = sheet.getRow(0);
        if (firstRow == null) {
            throw new IllegalStateException("The sheet is empty.");
        }

        for (int colIndex = 1; colIndex < firstRow.getLastCellNum(); colIndex++) {
            Cell cell = firstRow.getCell(colIndex);
            if (cell != null) {
                testCaseMap.put(colIndex, getCellValue(cell));
            }
        }

        for (int rowIndex = 1; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
            Row row = sheet.getRow(rowIndex);
            if (row == null) continue;

            Cell headerCell = row.getCell(0);
            if (headerCell == null) continue;

            String header = getCellValue(headerCell);
            if (!dataMap.containsKey(header)) {
                dataMap.put(header, new HashMap<>());
            }

            Map<String, String> testCaseValues = dataMap.get(header);
            for (Map.Entry<Integer, String> entry : testCaseMap.entrySet()) {
                int colIndex = entry.getKey();
                String testCase = entry.getValue();
                Cell cell = row.getCell(colIndex);
                testCaseValues.put(testCase, cell != null ? getCellValue(cell) : "");
            }
        }

        return dataMap;
    }

    public Map<String, Map<String, Map<String, String>>> readAllSheetsData() {
        Map<String, Map<String, Map<String, String>>> allData = new HashMap<>();
        for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
            Sheet sheet = workbook.getSheetAt(i);
            allData.put(sheet.getSheetName(), readSheetData(sheet.getSheetName()));
        }
        return allData;
    }

    public int getRowCount(String sheetName) {
        Sheet sheet = workbook.getSheet(sheetName);
        if (sheet == null) {
            throw new IllegalArgumentException("Sheet with name " + sheetName + " does not exist.");
        }
        return sheet.getPhysicalNumberOfRows();
    }

    public int getColCount(String sheetName) {
        Sheet sheet = workbook.getSheet(sheetName);
        if (sheet == null) {
            throw new IllegalArgumentException("Sheet with name " + sheetName + " does not exist.");
        }
        Row firstRow = sheet.getRow(0);
        if (firstRow == null) {
            return 0;
        }
        return firstRow.getLastCellNum();
    }

    private int getHeaderRowIndex(String sheetName, String headerName) {
        Sheet sheet = workbook.getSheet(sheetName);
        if (sheet == null) {
            throw new IllegalArgumentException("Sheet with name " + sheetName + " does not exist.");
        }

        for (int rowIndex = 0; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
            Row row = sheet.getRow(rowIndex);
            if (row != null) {
                Cell cell = row.getCell(0);
                if (cell != null && getCellValue(cell).equalsIgnoreCase(headerName)) {
                    return rowIndex;
                }
            }
        }
        throw new IllegalArgumentException("Header with name " + headerName + " does not exist.");
    }

    private int getTestCaseColIndex(String sheetName, String testCaseName) {
        Sheet sheet = workbook.getSheet(sheetName);
        if (sheet == null) {
            throw new IllegalArgumentException("Sheet with name " + sheetName + " does not exist.");
        }

        Row firstRow = sheet.getRow(0);
        if (firstRow != null) {
            for (int colIndex = 0; colIndex < firstRow.getLastCellNum(); colIndex++) {
                Cell cell = firstRow.getCell(colIndex);
                if (cell != null && getCellValue(cell).equalsIgnoreCase(testCaseName)) {
                    return colIndex;
                }
            }
        }
        throw new IllegalArgumentException("Test case with name " + testCaseName + " does not exist.");
    }

   public void updateExcel(String sheetName, String headerName, String testCaseName, String newValue, Map<String, Map<String, Map<String, String>>> dataMap) {
    Sheet sheet = workbook.getSheet(sheetName);
    if (sheet == null) {
        throw new IllegalArgumentException("Sheet with name " + sheetName + " does not exist.");
    }

    // Get row and column indices
    int headerRowIndex = getHeaderRowIndex(sheetName, headerName);
    int testCaseColIndex = getTestCaseColIndex(sheetName, testCaseName);

    // Update Excel sheet
    Row rowToUpdate = sheet.getRow(headerRowIndex);
    Cell cellToUpdate = rowToUpdate.getCell(testCaseColIndex);
    if (cellToUpdate == null) {
        cellToUpdate = rowToUpdate.createCell(testCaseColIndex);
    }
    cellToUpdate.setCellValue(newValue);

    // Update the in-memory data map
    if (!dataMap.containsKey(sheetName)) {
        dataMap.put(sheetName, new HashMap<>());
    }
    Map<String, Map<String, String>> sheetData = dataMap.get(sheetName);

    if (!sheetData.containsKey(headerName)) {
        sheetData.put(headerName, new HashMap<>());
    }
    Map<String, String> testCaseMap = sheetData.get(headerName);
    testCaseMap.put(testCaseName, newValue);

    // Save changes to the file
    try (FileOutputStream fos = new FileOutputStream(filePath)) {
        workbook.write(fos);
    } catch (IOException e) {
        throw new RuntimeException("Failed to write to the Excel file.", e);
    }
}

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

    public void close() throws IOException {
        if (workbook != null) {
            workbook.close();
        }
    }

    public static void main(String[] args) {
        try {
            ExcelUtil excelUtil = new ExcelUtil("path/to/your/excel/file.xlsx");

            Map<String, Map<String, String>> sheetData = excelUtil.readSheetData("Sheet1");
            System.out.println("Sheet1 Data: " + sheetData);

            Map<String, Map<String, Map<String, String>>> allData = excelUtil.readAllSheetsData();
            System.out.println("All Sheets Data: " + allData);

            System.out.println("Row Count: " + excelUtil.getRowCount("Sheet1"));
            System.out.println("Column Count: " + excelUtil.getColCount("Sheet1"));

            excelUtil.updateExcel("Sheet1", "Header1", "TestCase1", "NewValue");

            excelUtil.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}

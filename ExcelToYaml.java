import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.yaml.snakeyaml.DumperOptions;
import org.yaml.snakeyaml.Yaml;
import org.testng.xml.XmlClass;
import org.testng.xml.XmlInclude;
import org.testng.xml.XmlSuite;
import org.testng.xml.XmlTest;
import org.testng.xml.XmlMethodSelector;

import java.io.*;
import java.util.*;

public class ExcelToYaml {

    public static void main(String[] args) throws Exception {
        // Specify the path of the input Excel file and output YAML file and TestNG XML file
    	String excelFilePath = "C:\\baxterCode\\TestData.xlsx";
        String yamlFilePath = "C:\\baxterCode\\output.yaml";
        String testngXmlFilePath = "C:\\baxterCode\\output-testng.xml";

        // Create the YAML file and TestNG XML file from the Excel data
        Workbook workbook = new XSSFWorkbook(new FileInputStream(excelFilePath));
        createYamlFile(workbook, yamlFilePath);
        createTestNgXml(workbook, testngXmlFilePath);
        workbook.close();
    }

    private static void createYamlFile(Workbook workbook, String yamlFilePath) throws IOException {
        Map<String, Object> yamlData = new LinkedHashMap<>();
        
        // Iterate through each sheet
        for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
            Sheet sheet = workbook.getSheetAt(i);
            String environmentName = sheet.getSheetName();
            
            // Skip "Execution" sheet
            if ("Execution".equalsIgnoreCase(environmentName)) {
                continue;
            }

            Map<String, Object> envData = new LinkedHashMap<>();
            Map<String, Object> testCases = new LinkedHashMap<>();

            // Iterate through each row in the sheet
            for (int rowIndex = 1; rowIndex <= sheet.getPhysicalNumberOfRows(); rowIndex++) {
                Row row = sheet.getRow(rowIndex);
                if (row == null) continue;

                // Read cells for TestCases, OQName, StepNo, FieldName, and Value
                Cell testCaseCell = row.getCell(0);
                Cell oqNameCell = row.getCell(1);
                Cell stepNoCell = row.getCell(2);
                Cell fieldNameCell = row.getCell(3);
                Cell valueCell = row.getCell(4);

                // Check for non-null and non-blank cells to handle each case accordingly
                String testCase = (testCaseCell != null) ? testCaseCell.getStringCellValue().trim() : "";
                String oqName = (oqNameCell != null) ? oqNameCell.getStringCellValue().trim() : "";
                String stepNo = (stepNoCell != null) ? stepNoCell.getStringCellValue().trim() : "";
                String fieldName = (fieldNameCell != null) ? fieldNameCell.getStringCellValue().trim() : "";
                String value = (valueCell != null) ? valueCell.getStringCellValue().trim() : "";

                // Handle the cases based on columns
                if (!testCase.isEmpty() && !oqName.isEmpty() && !stepNo.isEmpty() && !fieldName.isEmpty()) {
                    // TestCase -> OQName -> StepNo -> FieldName
                    if (!testCases.containsKey(testCase)) {
                        testCases.put(testCase, new LinkedHashMap<>());
                    }
                    Map<String, Object> oqData = (Map<String, Object>) testCases.get(testCase);
                    if (!oqData.containsKey(oqName)) {
                        oqData.put(oqName, new LinkedHashMap<>());
                    }
                    Map<String, Object> stepData = (Map<String, Object>) oqData.get(oqName);
                    if (!stepData.containsKey(stepNo)) {
                        stepData.put(stepNo, new LinkedHashMap<>());
                    }
                    Map<String, String> fieldData = (Map<String, String>) stepData.get(stepNo);
                    fieldData.put(fieldName, value);
                } else if (!testCase.isEmpty() && !oqName.isEmpty() && !fieldName.isEmpty()) {
                    // TestCase -> OQName -> FieldName
                    if (!testCases.containsKey(testCase)) {
                        testCases.put(testCase, new LinkedHashMap<>());
                    }
                    Map<String, Object> oqData = (Map<String, Object>) testCases.get(testCase);
                    if (!oqData.containsKey(oqName)) {
                        oqData.put(oqName, new LinkedHashMap<>());
                    }
                    Map<String, String> fieldData = (Map<String, String>) oqData.get(oqName);
                    fieldData.put(fieldName, value);
                } else if (!fieldName.isEmpty()) {
                    // Just FieldName
                    envData.put(fieldName, value);
                } else if (!testCase.isEmpty() && !fieldName.isEmpty()) {
                    // TestCase -> FieldName
                    if (!testCases.containsKey(testCase)) {
                        testCases.put(testCase, new LinkedHashMap<>());
                    }
                    Map<String, String> fieldData = (Map<String, String>) testCases.get(testCase);
                    fieldData.put(fieldName, value);
                }
            }

            // Add test cases and environment data to the main YAML data
            if (!testCases.isEmpty()) {
                envData.put("TestCases", testCases.get("TestCases"));
            }
            yamlData.put(environmentName, envData);
        }

        // Write the data to a YAML file (using SnakeYAML for example)
        try (FileWriter writer = new FileWriter(yamlFilePath)) {
            DumperOptions options = new DumperOptions();
            options.setIndent(4);
            options.setDefaultFlowStyle(DumperOptions.FlowStyle.BLOCK);
            Yaml yaml = new Yaml(options);
            yaml.dump(yamlData, writer);
        }
    }

    private static void createTestNgXml(Workbook workbook, String testngXmlFilePath) throws IOException {
    	XmlSuite suite = new XmlSuite();
        suite.setName("Suite");
        suite.setParallel(XmlSuite.ParallelMode.METHODS); // Set parallel execution for methods
        suite.setThreadCount(1);

        XmlTest test = new XmlTest(suite);
        test.setName("BHS Functional Test");

        // Add listeners to the suite
        List<String> listeners = new ArrayList<>();
        listeners.add("com.report.Listners");
        suite.setListeners(listeners);

        // Add the test class
        XmlClass testClass = new XmlClass("com.bhs.tests.BHS_FunctionalTests");
        test.getClasses().add(testClass);

        // Get the "Execution" sheet to determine which tests to include
        Sheet executionSheet = workbook.getSheet("Execution");
        if (executionSheet != null) {
            for (int rowIndex = 1; rowIndex <= executionSheet.getPhysicalNumberOfRows(); rowIndex++) {
                Row row = executionSheet.getRow(rowIndex);
                if (row == null) continue;

                // Read the OQName and Execute cells
                Cell oqNameCell = row.getCell(0);
                Cell executeCell = row.getCell(1);

                if (oqNameCell != null && executeCell != null) {
                    String oqName = oqNameCell.getStringCellValue().trim();
                    String execute = executeCell.getStringCellValue().trim();

                    // Include methods where Execute is Yes or Y (case-insensitive)
                    if ("Y".equalsIgnoreCase(execute) || "Yes".equalsIgnoreCase(execute)) {
                        testClass.getIncludedMethods().add(new XmlInclude(oqName));
                    }
                }
            }
        }

        // Write the TestNG XML configuration to file
        try (FileOutputStream outputStream = new FileOutputStream(testngXmlFilePath)) {
        	String xmlContent = suite.toXml(); // Get the XML string
            outputStream.write(xmlContent.getBytes()); // Write the string to the file
        }
    }
}

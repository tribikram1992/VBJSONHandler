import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.yaml.snakeyaml.DumperOptions;
import org.yaml.snakeyaml.Yaml;

import java.io.*;
import java.util.*;

public class YamlToExcel {

	public static void main(String[] args) throws Exception {
		// Specify the path of the input YAML file and output Excel file
		String yamlFilePath = "C:\\baxterCode\\output.yaml";
		String excelFilePath = "C:\\baxterCode\\TestData1.xlsx";

		// Read the YAML file and convert it to Excel
		Map<String, Object> yamlData = readYamlFile(yamlFilePath);
		createExcelFile(yamlData, excelFilePath);
	}

	private static Map<String, Object> readYamlFile(String yamlFilePath) throws IOException {
		Yaml yaml = new Yaml();
		try (FileReader reader = new FileReader(yamlFilePath)) {
			return yaml.load(reader);
		}
	}

	private static void createExcelFile(Map<String, Object> yamlData, String excelFilePath) throws IOException {
		Workbook workbook = new XSSFWorkbook();

		// Create a sheet for each environment
		for (Map.Entry<String, Object> entry : yamlData.entrySet()) {
			String environmentName = entry.getKey();
			Object environmentData = entry.getValue();

			// Create a new sheet for the environment
			Sheet sheet = workbook.createSheet(environmentName);
			int rowIndex = 0;

			// Set up the header row with required columns: TestCases, OQName, StepNo,
			// FieldName, Value
			Row headerRow = sheet.createRow(rowIndex++);
			headerRow.createCell(0).setCellValue("TestCases");
			headerRow.createCell(1).setCellValue("OQName");
			headerRow.createCell(2).setCellValue("StepNo");
			headerRow.createCell(3).setCellValue("FieldName");
			headerRow.createCell(4).setCellValue("Value");

			// Now process the environment data
			if (environmentData instanceof Map) {
				Map<String, Object> environmentMap = (Map<String, Object>) environmentData;

				for(Map.Entry<String, Object> environmentEntry : environmentMap.entrySet()) {
					String environmentkey = environmentEntry.getKey();
					Object environmentvalue = environmentEntry.getValue();
					
					if(environmentvalue instanceof Map) {
						// Process TestCases if present
						if (environmentMap.containsKey("TestCases")) {
							Map<String, Object> testCases = (Map<String, Object>) environmentMap.get("TestCases");

							for (Map.Entry<String, Object> testCaseEntry : testCases.entrySet()) {
								String oqName = testCaseEntry.getKey();
								Object testCaseData = testCaseEntry.getValue();

								// Process OQName and StepNo if present
								if (testCaseData instanceof Map) {
									Map<String, Object> testCaseMap = (Map<String, Object>) testCaseData;

									for (Map.Entry<String, Object> oqNameEntry : testCaseMap.entrySet()) {
										String stepNo = oqNameEntry.getKey();
										Object stepData = oqNameEntry.getValue();

										if (stepData instanceof Map) {
											Map<String, Object> stepMap = (Map<String, Object>) stepData;
											for (Map.Entry<String, Object> stepEntry : stepMap.entrySet()) {
												String fieldName = stepEntry.getKey();
												String value = (String) stepEntry.getValue();
												Row row = sheet.createRow(rowIndex++);
												row.createCell(0).setCellValue("TestCases");
												row.createCell(1).setCellValue(oqName);
												row.createCell(2).setCellValue(stepNo);
												row.createCell(3).setCellValue(fieldName);
												row.createCell(4).setCellValue(value);
											}
										}
										else {
											Row row = sheet.createRow(rowIndex++);
											row.createCell(0).setCellValue("TestCases");
											row.createCell(1).setCellValue(oqName);
											row.createCell(2).setCellValue("");
											row.createCell(3).setCellValue(stepNo);
											row.createCell(4).setCellValue((String)stepData);
											
										}

									}
								}
							}
						} 

						}
					else {
						Row row = sheet.createRow(rowIndex++);
						row.createCell(0).setCellValue("");
						row.createCell(1).setCellValue("");
						row.createCell(2).setCellValue("");
						row.createCell(3).setCellValue(environmentkey);
						row.createCell(4).setCellValue((String)environmentvalue);
					}
				}
			}
		}

		// Write the workbook to an Excel file
		try (FileOutputStream fileOut = new FileOutputStream(excelFilePath)) {
			workbook.write(fileOut);
		}

		workbook.close();
	}
}

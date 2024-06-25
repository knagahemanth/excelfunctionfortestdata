package com.util;

import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.io.File;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelFunctions {

	public LinkedHashMap<String, String> readTestData(String sTcId) throws Exception {
		List<String> tcList = getValuesInColumn(System.getProperty("user.dir"), "Companytestdata", "Sheet2", 0);
		int rowCount = 0;
		for (int count = 0; count < tcList.size(); count++) {
			if (tcList.get(count).equals(sTcId)) {
				rowCount = count;
				break;
			}

		}
		return getRowDataHM(System.getProperty("user.dir"), "Companytestdata", "Sheet2", rowCount+1);
	}

	// For Excel Interaction
	public LinkedHashMap<String, String> getRowDataHM(String FilePath, String WorkBookName, String SheetName,
			int rowIndex) throws Exception {

		FileInputStream file = new FileInputStream(new File(FilePath + "\\" + WorkBookName + ".xlsx"));
		XSSFWorkbook book = new XSSFWorkbook(file);
		XSSFSheet sheet = book.getSheet(SheetName);
		// System.out.println("sheet"+SheetName);
		XSSFRow row = sheet.getRow(rowIndex);
		XSSFRow headerRow = sheet.getRow(0);
		LinkedHashMap<String, String> data = new LinkedHashMap<String, String>();
		int firstCell = headerRow.getFirstCellNum();
		int lastCell = headerRow.getLastCellNum();
		XSSFCell cell1 = headerRow.getCell(firstCell);
		XSSFCell cell2 = row.getCell(firstCell);

		for (int i = firstCell + 1; i < lastCell; i++) {
			cell1 = headerRow.getCell(i);
			cell2 = row.getCell(i);
			String headerValue = cell1.getStringCellValue();
			String fieldvalue;
			if (cell2 == null) {
				fieldvalue = "";
			} else {
				fieldvalue = cell2.getStringCellValue();
			}
			data.put(headerValue, fieldvalue);

		}
		book.close();
		return data;
	}

	public static int getRows(String FilePath, String WorkBookName, String SheetName) throws Exception {
		FileInputStream file = new FileInputStream(new File(FilePath + "\\" + WorkBookName + ".xlsx"));
		XSSFWorkbook book = new XSSFWorkbook(file);
		XSSFSheet sheet = book.getSheet(SheetName);
		int totalRows = sheet.getPhysicalNumberOfRows();
		book.close();
		return totalRows;
	}

	public static List<String> getValuesInColumn(String FilePath, String WorkBookNambe, String SheetName, int colIndex)
			throws Exception {
		FileInputStream file = new FileInputStream(new File(FilePath + "\\" + WorkBookName + ".xlsx"));
		XSSFWorkbook book = new XSSFWorkbook(file);
		XSSFSheet sheet = book.getSheet(SheetName);
		int totalRows = sheet.getPhysicalNumberOfRows();
		List<String> data = new ArrayList<String>();
		for (int i = 1; i < totalRows; i++) {
			XSSFRow row = sheet.getRow(i);
			int cellNo = colIndex;
			XSSFCell cell = row.getCell(cellNo);
			if (cell == null) {
				continue;
			} else {
				data.add(cell.getStringCellValue());
			}
		}
		book.close();
		return data;
	}

	public static List<String> getValuesInColumn(String FilePath, String WorkBookName, String SheetName, String colName)
			throws Exception {
		FileInputStream file = new FileInputStream(new File(FilePath + "\\" + WorkBookName + ".xlsx"));
		XSSFWorkbook book = new XSSFWorkbook(file);
		XSSFSheet sheet = book.getSheet(SheetName);
		int totalRows = sheet.getPhysicalNumberOfRows();
		List<String> data = new ArrayList<String>();
		int colIndex = getColumnWithName(FilePath, WorkBookName, SheetName, colName);
		for (int i = 1; i < totalRows; i++) {
			XSSFRow row = sheet.getRow(i);
			int cellNo = colIndex;
			XSSFCell cell = row.getCell(cellNo);
			if (cell == null) {
				continue;
			} else {
				data.add(cell.getStringCellValue());
			}
		}
		book.close();
		return data;
	}

	public List<String> getValuesForcolumnwithTestDataID(String FilePath, String WorkBookName, String SheetName,
			String colName, String testDataID) {
		try {
			FileInputStream file = new FileInputStream(new File(FilePath + "\\" + WorkBookName + ".xlsx"));
			XSSFWorkbook book = new XSSFWorkbook(file);
			XSSFSheet sheet = book.getSheet(SheetName);
			int totalRows = sheet.getPhysicalNumberOfRows();
			List<String> data = new ArrayList<String>();
			int colIndex = getColumnWithName(FilePath, WorkBookName, SheetName, colName);
			for (int i = 1; i < totalRows; i++) {
				XSSFRow row = sheet.getRow(i);
				XSSFCell testDataIDCell = row.getCell(1);

				if (testDataIDCell.getStringCellValue().equalsIgnoreCase(testDataID)) {
					XSSFCell cell = row.getCell(colIndex);
					if ((cell == null) || (cell.getStringCellValue().equalsIgnoreCase("NA"))) {
						continue;
					} else {
						String cellVal = cell.getStringCellValue();
						String[] values = cellVal.split(",");
						for (String val : values) {
							data.add(val);
						}
					}
				}
			}
			book.close();
			return data;
		} catch (Exception e) {
			return null;
		}
	}

	public static int getColumnWithName(String FilePath, String WorkBookName, String SheetName, String columnName)
			throws Exception {
		FileInputStream file = new FileInputStream(new File(FilePath + "\\" + WorkBookName + ".xlsx"));
		XSSFWorkbook book = new XSSFWorkbook(file);
		XSSFSheet sheet = book.getSheet(SheetName);
		XSSFRow row = sheet.getRow(0);
		int totalCol = row.getPhysicalNumberOfCells();
		int retVal = 0;
		boolean isFound = false;
		for (int i = 0; i < totalCol; i++) {
			XSSFCell cell = row.getCell(i);
			if (cell == null) {
				continue;
			} else {
				if (cell.getStringCellValue().trim().equalsIgnoreCase(columnName.trim())) {
					book.close();
					retVal = i;
					isFound = true;
				}
			}
		}
		if (isFound) {
			book.close();
			return retVal;
		} else {
			throw new Exception(WorkBookName);
		}
	}

	// To handle nun the same Scenario multiple times wAth different test data rows
	public List<LinkedHashMap<String, String>> getTestDataListWithScenario(String FilePath, String WorkBookName,
			String SheetName, String scenarioName) throws Exception {
		FileInputStream file = new FileInputStream(new File(FilePath + "\\" + WorkBookName + ".xlsx"));
		XSSFWorkbook book = new XSSFWorkbook(file);
		XSSFSheet sheet = book.getSheet(SheetName);
		List<LinkedHashMap<String, String>> data = new ArrayList<LinkedHashMap<String, String>>();
		int totalRows = sheet.getPhysicalNumberOfRows();
		for (int i = 1; i < totalRows; i++) {
			XSSFRow row = sheet.getRow(i);
			XSSFCell cell = row.getCell(0);
			try {
				if (cell.getStringCellValue().equalsIgnoreCase(scenarioName)) {
					data.add(getRowDataHM(FilePath, WorkBookName, SheetName, i));
				}
			} catch (Exception e) {
				break;
			}
		}
		book.close();
		return data;
	}

	public LinkedHashMap<String, String> getRowDataHMScenario(String FilePath, String WorkBookName, String SheetName,
			int rowIndex, String scenarioName) throws Exception {
		FileInputStream file = new FileInputStream(new File(FilePath + "\\" + WorkBookName + ". xlsx"));
		XSSFWorkbook book = new XSSFWorkbook(file);
		XSSFSheet sheet = book.getSheet(SheetName);
		XSSFRow row = sheet.getRow(rowIndex);
		XSSFRow headerRow = sheet.getRow(0);

		LinkedHashMap<String, String> data = new LinkedHashMap<String, String>();
		int firstCell = headerRow.getFirstCellNum();
		int lastCell = headerRow.getLastCellNum();
		int totalRows = sheet.getPhysicalNumberOfRows();
		XSSFCell cell1 = headerRow.getCell(firstCell);
		XSSFCell cell2 = row.getCell(firstCell);
		for (int i = 1; i < totalRows; i++) {
			XSSFRow row1 = sheet.getRow(i);
			XSSFCell cell = row.getCell(0);
			if (cell.getStringCellValue().equalsIgnoreCase(scenarioName)) {
				for (i = firstCell; i < lastCell; i++) {
					cell1 = headerRow.getCell(i);
					cell2 = row.getCell(i);
					String headerValue = cell1.getStringCellValue();
					String fieldValue;
					if (cell2 == null) {
						fieldValue = "";
					} else {
						fieldValue = cell2.getStringCellValue();
					}
					data.put(headerValue, fieldValue);
				}
				break;
			}
			break;
		}
		book.close();
		return data;
	}

	public static void main(String[] args) throws Exception {
		ExcelFunctions gf = new ExcelFunctions();
	}
}

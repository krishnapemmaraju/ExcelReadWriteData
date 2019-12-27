package com.test.qa.utils;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class GetExcelData {

	public static String path;
	public static FileInputStream fis;
	public static XSSFWorkbook workbook;
	public static XSSFSheet sheet;
	public static XSSFRow row;
	public static XSSFCell cell;
	public static FileOutputStream fos;

	public GetExcelData(String path) {
		this.path = path;
		try {
			fis = new FileInputStream(path);
			workbook = new XSSFWorkbook(fis);
			sheet = workbook.getSheetAt(0);
			fis.close();
		} catch (Exception e) {

		}
	}

	/*
	 * Getting the Row Count Author - Krishna Pemmaraju
	 */

	public static int getRowCount(String sheetName) {
		int index = workbook.getSheetIndex(sheetName);
		if (index == -1)
			return 0;
		else {
			return sheet.getLastRowNum() + 1;
		}
	}

	/*
	 * Adding New Sheet Author - Krishna Pemmaraju
	 */

	public static boolean createSheet(String sheetName) {
		FileOutputStream fos;
		try {
			workbook.createSheet(sheetName);
			fos = new FileOutputStream(path);
			workbook.write(fos);
			fos.close();
		} catch (Exception e) {
			e.printStackTrace();
			return false;
		}
		return true;
	}

	public static boolean removeSheet(String sheetName) {
		int index = workbook.getSheetIndex(sheetName);
		FileOutputStream fos;
		if (index == -1)
			return false;
		try {
			workbook.removeSheetAt(index);
			fos = new FileOutputStream(path);
			workbook.write(fos);
			return true;
		} catch (Exception e) {
			e.printStackTrace();
			return false;
		}
	}

	public static boolean isSheetExists(String sheetName) {
		int index = workbook.getSheetIndex(sheetName);
		if (index == -1) {
			return false;
		} else {
			return true;
		}
	}

	public static int getColumnCount(String sheetName) {
		int index = workbook.getSheetIndex(sheetName);
		if (index == -1)
			return -1;
		else {
			sheet = workbook.getSheet(sheetName);
			return sheet.getRow(0).getLastCellNum();
		}
	}

	public static String getCellRowNum(String sheetName, String colName, String cellValue) {
		sheet = workbook.getSheet(sheetName);
		for (int i = 0; i <= sheet.getLastRowNum(); i++) {
			for (int j = 0; j < sheet.getRow(0).getLastCellNum(); j++) {
				System.out.println(sheet.getRow(i).getCell(j));
				if (sheet.getRow(i).getCell(j).toString().equalsIgnoreCase(cellValue))
					return " Row number is  " + i + " Column Number is " + j;
			}
		}
		return "No Value Found";
	}

	public static boolean removeColNum(String sheetName, int ColNum) throws IOException {
		int index = workbook.getSheetIndex(sheetName);
		if (index == -1)
			return false;
		fis = new FileInputStream(path);
		sheet = workbook.getSheet(sheetName);
		for (int i = 0; i < sheet.getLastRowNum(); i++) {
			row = sheet.getRow(i);
			if (row != null) {
				cell = row.getCell(ColNum);
				if (cell != null) {
					row.removeCell(cell);
				}
			}
		}
		fos = new FileOutputStream(path);
		workbook.write(fos);
		fos.close();
		return true;
	}

	public static String getCellData(String sheetName, String colName, int rowNum) {
		int colNum = -1;
		if (rowNum == -1)
			return "";

		int index = workbook.getSheetIndex(sheetName);
		if (index == -1)
			return "";

		sheet = workbook.getSheet(sheetName);
		row = sheet.getRow(0);
		for (int i = 0; i < sheet.getLastRowNum(); i++) {
			if (row.getCell(i).getStringCellValue().trim().equalsIgnoreCase(colName))
				colNum = i;
		}
		if (colNum == -1)
			return "";
		sheet = workbook.getSheetAt(index);
		row = sheet.getRow(rowNum + 1);
		if (row == null)
			return "";
		cell = row.getCell(colNum);
		return cell.getStringCellValue();
	}

	public static boolean createNewColumn(String sheetName, String colName) throws IOException {
		int colNum = -1;
		int index = workbook.getSheetIndex(sheetName);
		if (index == -1)
			return false;

		row = sheet.getRow(0);
		colNum = sheet.getRow(0).getLastCellNum();
		System.out.println("The no of Columns in First Row is " + colNum);
		if (row == null)
			return false;
		for (int i = 0; i <= colNum; i++) {
			System.out.println("Tha value is " + sheet.getRow(0).getCell(i));
			if (sheet.getRow(0).getCell(i) == null) {
				cell = row.createCell(i);
				cell.setCellValue(colName);
			}
		}
		fos = new FileOutputStream(path);
		workbook.write(fos);
		fos.close();
		return true;
	}

	public static String setCellData(String sheetName, String data, String colName, int rowNum) throws IOException {

		fis = new FileInputStream(path);
		int index = workbook.getSheetIndex(sheetName);
		int colNum = -1;
		if (index == -1)
			return "No Sheet";
		row = sheet.getRow(0);
		for (int i = 0; i < row.getLastCellNum(); i++) {
			if (row.getCell(i).getStringCellValue().equalsIgnoreCase(colName)) {
				colNum = i;
				System.out.println(row.getCell(i).toString());
				System.out.println("The ColNum is " + i);
			}
		}
		sheet.autoSizeColumn(colNum);
		if (colNum == -1)
			return "No Col Value";
		System.out.println("The Row Num to create a row is " + rowNum);
		row = sheet.getRow(rowNum - 1);
		if (row == null) {
			System.out.println("The Row Num to create a row inside is " + rowNum);
		row = sheet.createRow(rowNum - 1);}
		cell = row.getCell(colNum);
		if (cell == null)
			cell = row.createCell(colNum);
		cell.setCellValue(data);
		fos = new FileOutputStream(path);
		workbook.write(fos);
		fos.close();
		return "Data Inserted";
	}

}

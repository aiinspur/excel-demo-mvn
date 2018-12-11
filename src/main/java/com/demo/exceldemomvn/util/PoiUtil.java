package com.demo.exceldemomvn.util;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class PoiUtil {

	public static List<Sheet> getSheetList(File file) throws FileNotFoundException, IOException {
		List<Sheet> sheets = null;

		try (InputStream inp = new FileInputStream(file)) {
			Workbook wb = WorkbookFactory.create(inp);
			int numberOfSheets = wb.getNumberOfSheets();
			sheets = new ArrayList<>(numberOfSheets);
			for (int index = 0; index < numberOfSheets; index++) {
				sheets.add(wb.getSheetAt(index));
			}
		}
		return sheets;
	}

	public static List<String> getSheetNameList(File file) throws FileNotFoundException, IOException {
		List<Sheet> sheetList = getSheetList(file);
		List<String> sheetNameList = new ArrayList<>(sheetList.size());
		sheetList.forEach(sheet -> {
			sheetNameList.add(sheet.getSheetName().trim());
			System.out.println(sheet.getSheetName().trim());
		});
		return sheetNameList;
	}

	public static void copySheetFromSheet(Sheet srcSheet, Sheet destSheet, File destFile,Workbook destWb) throws Exception {
//		try (InputStream inp = new FileInputStream(destFile)) {
//			Workbook wb = WorkbookFactory.create(inp);
//
//			copyRowData(wb, destSheet, srcSheet);
//
//			FileOutputStream fileOut = new FileOutputStream(destFile);
//			wb.write(fileOut);
//		}
		copyRowData( destSheet, srcSheet);

		FileOutputStream fileOut = new FileOutputStream(destFile);
		destWb.write(fileOut);
	}

	public static void createSheet(Sheet sheet, File destFile) throws Exception {

		if (!destFile.exists()) {
			try (HSSFWorkbook wb = new HSSFWorkbook()) {
				wb.createSheet(sheet.getSheetName());

				try (FileOutputStream fileOut = new FileOutputStream(destFile)) {
					wb.write(fileOut);
				}
			}
		} else {
			try (InputStream inp = new FileInputStream(destFile)) {
				Workbook wb = WorkbookFactory.create(inp);
				wb.createSheet(sheet.getSheetName());

				FileOutputStream fileOut = new FileOutputStream(destFile);
				wb.write(fileOut);
			}

		}

	}

	public static void createSheets(File excelFile, List<Sheet> sheets) throws Exception {
		sheets.forEach(sheet -> {
			try {
				createSheet(sheet, excelFile);
			} catch (Exception e) {
				e.printStackTrace();
			}
		});
	}

	public static void copyRowData(Sheet destSheet, Sheet srcSheet) {
		int firstRowNum = srcSheet.getFirstRowNum();
		int lastRowNum = srcSheet.getLastRowNum();

		for (int i = firstRowNum; i < lastRowNum; i++) {
			Row srcRow = srcSheet.getRow(i);
			Row destRow = destSheet.getRow(i);

			if (srcRow == null || destRow == null) {
				continue;
			}

			short minColIx = srcRow.getFirstCellNum();
			short maxColIx = srcRow.getLastCellNum();
			for (short j = minColIx; j < maxColIx; j++) {
				Cell cell = srcRow.getCell(j);

				Cell destCell = destRow.getCell(j);
				if (cell == null || destCell == null) {
					continue;
				}

				CellType cellType = cell.getCellType();
				switch (cellType) {
				case STRING:
					destCell.setCellValue(cell.getStringCellValue());
					break;
				case NUMERIC:
					destCell.setCellValue(cell.getNumericCellValue());
					break;
				case FORMULA:
					destCell.setCellValue(cell.getCellFormula());
					break;
				case BOOLEAN:
					destCell.setCellValue(cell.getBooleanCellValue());
					break;
				case ERROR:
					destCell.setCellValue(cell.getErrorCellValue());
					break;
				default:
					destCell.setCellValue(cell.getStringCellValue());
					break;
				}
			}
			
			
		}
	}

}

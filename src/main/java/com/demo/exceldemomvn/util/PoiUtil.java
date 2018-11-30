package com.demo.exceldemomvn.util;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.springframework.util.ResourceUtils;

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
			sheetNameList.add(sheet.getSheetName());
		});
		return sheetNameList;
	}

}

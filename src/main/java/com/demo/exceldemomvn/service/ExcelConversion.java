package com.demo.exceldemomvn.service;

import com.demo.exceldemomvn.util.PoiUtil;
import org.apache.poi.ss.usermodel.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.core.io.ClassPathResource;
import org.springframework.stereotype.Service;

import java.io.*;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

@Service
public class ExcelConversion implements FileConversion {

	private final Logger logger = LoggerFactory.getLogger(ExcelConversion.class);

	@Override
	public void conversion(String srcFile, String destinationFilePath, String dataDate) throws Exception {
		logger.info("src file:" + srcFile + ";destinationFilePath:" + destinationFilePath);

		if (!new File(destinationFilePath).exists()) {
			new File(destinationFilePath).mkdirs();
		}

		String destFile = destinationFilePath + File.separator + "提供行领导" + dataDate + ".xls";

		try (InputStream inp = new ClassPathResource("excel_template/template_1.xls").getInputStream()) {
			Workbook wb = WorkbookFactory.create(inp);
			FileOutputStream fileOut = new FileOutputStream(destFile);
			wb.write(fileOut);
		}

		Map<String, String> map = getMapping2(dataDate);

		Workbook srcWb = null;
		try (InputStream inp = new FileInputStream(srcFile)) {
			srcWb = WorkbookFactory.create(inp);
		}

		Workbook destWb = null;
		try (InputStream inp = new FileInputStream(destFile)) {
			destWb = WorkbookFactory.create(inp);
		}

		
		for (String key : map.keySet()) {
			String val = map.get(key);
			if (val.equals("")) {
				continue;
			}
			//PoiUtil.copySheetFromSheet(srcWb.getSheet(val), destWb.getSheet(key), new File(destFile), destWb);
			
			FileOutputStream fileOut = new FileOutputStream(destFile);
			Sheet destSheet = destWb.getSheet(key);
			copyRowData(destSheet, srcWb.getSheet(val));
			destWb.write(fileOut);
			fileOut.flush();
			System.out.println("---------"+destSheet.getSheetName());
			fileOut.close();
		}
		
		if (srcWb != null) {
			srcWb.close();
		}
		
		if (destWb != null) {
			destWb.close();
		}

	}
	
	public void copyRowData(Sheet destSheet, Sheet srcSheet) {
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

	/**
	 * 手动维护映射关系
	 * 
	 * @return
	 * @throws Exception
	 */
	private Map<String, String> getMapping2(String srcDate) throws Exception {
		Map<String, String> map = new LinkedHashMap<>();
		// src
		/*
		 * R0008-2017_全行_20181120 R0009-2017_全行_20181120 R0010-2017_全行_20181120
		 * R0011-2017_全行_20181120 R0012-2017_全行_20181120 R0013-2017_全行_20181120
		 * R0014-2017_全行_20181120 R0016-2017_全行_20181120 R0017-2017_全行_20181120
		 * R0019-2017_全行_20181120 R0021-2017_全行_20181120 R0025-2017_全行_20181120
		 * R0030-2017_全行_20181120 R0040_全行_20181120 R0041_全行_20181120
		 * R0061N-2017_全行_20181120 R0061-2017_全行_20181120 R0062N-2017_全行_20181120
		 * R0062-2017_全行_20181120
		 */

		// template
		//
		/*
		 * 目录 R0030-2017_全行 R0041 主动负债 R0040_全行 R0061N-2017_全行 R0062N-2017_全行
		 * R0061-2017_全行 R0062-2017_全行 R0008 R0009 R0010 R0011 R0012 R0013 R0014 R0016
		 * R0021
		 */

		map.put("目录", "");
		map.put("R0030-2017_全行", "R0030-2017_全行_" + srcDate);
		map.put("R0041", "R0041_全行_" + srcDate);
		map.put("主动负债", "");
		map.put("R0040_全行", "R0040_全行_" + srcDate);
		map.put("R0061N-2017_全行", "R0061N-2017_全行_" + srcDate);
		map.put("R0062N-2017_全行", "R0062N-2017_全行_" + srcDate);
		map.put("R0061-2017_全行", "R0061-2017_全行_" + srcDate);
		map.put("R0062-2017_全行", "R0062-2017_全行_" + srcDate);
		map.put("R0008", "R0008-2017_全行_" + srcDate);
		map.put("R0009", "R0009-2017_全行_" + srcDate);
		map.put("R0010", "R0010-2017_全行_" + srcDate);
		map.put("R0011", "R0011-2017_全行_" + srcDate);
		map.put("R0012", "R0012-2017_全行_" + srcDate);
		map.put("R0013", "R0013-2017_全行_" + srcDate);
		map.put("R0014", "R0014-2017_全行_" + srcDate);
		map.put("R0016", "R0016-2017_全行_" + srcDate);
		map.put("R0021", "R0021-2017_全行_" + srcDate);

		return map;

	}

	/**
	 * 获取源文件与模版文件的sheet映射关系. 取交集
	 * 
	 * @param srcFile
	 * @param templateFile
	 * @return
	 * @throws IOException
	 * @throws FileNotFoundException
	 */
	@SuppressWarnings("unused")
	private Map<String, String> getMapping(File srcFile, File templateFile) throws Exception {
		Map<String, String> map = null;

		List<String> templateSheetNameList = PoiUtil.getSheetNameList(templateFile);
		// 获取转换后的模版sheet页
		List<String> templateSetList = conversionSheetName(templateSheetNameList);

		List<String> srcSheetNameList = PoiUtil.getSheetNameList(srcFile);
		// 获取转换后的源文件sheet页
		List<String> srcSetList = conversionSheetName(srcSheetNameList);
		// 取交集
		templateSetList.retainAll(srcSetList);

		// map = new LinkHashMap<>(templateSetList.size());
		map = new LinkedHashMap<>(templateSetList.size());

		for (String shortSheetName : templateSetList) {
			map.put(getSheetNameByShort(shortSheetName, templateSheetNameList),
					getSheetNameByShort(shortSheetName, srcSheetNameList));
		}

		return map;
	}

	private List<String> conversionSheetName(List<String> sheetNameList) {
		List<String> setList = new ArrayList<>(sheetNameList.size());
		sheetNameList.forEach(sheetName -> {
			sheetName = sheetName.trim();
			setList.add(sheetName.split("-")[0]);

//			if (sheetName.length() == 5) {
//				setList.add(sheetName.substring(0, 5));
//			} else if (sheetName.length() >= 6) {
//				setList.add(sheetName.substring(0, 6));
//			} else {
//				logger.info("sheetName error." + sheetName);
//			}
		});

		return setList;
	}

	private String getSheetNameByShort(String shortName, List<String> sheetNameList) {
		for (String name : sheetNameList) {
			if (name.startsWith(shortName)) {
				return name;
			}
		}
		return "";
	}

}

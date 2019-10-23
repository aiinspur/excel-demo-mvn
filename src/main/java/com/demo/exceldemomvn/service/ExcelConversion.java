package com.demo.exceldemomvn.service;

import com.demo.exceldemomvn.util.PoiUtil;
import org.apache.poi.ss.usermodel.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Service;

import java.io.*;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.*;

@Service
public class ExcelConversion implements FileConversion {

	private final Logger logger = LoggerFactory.getLogger(ExcelConversion.class);

	@Override
	public void conversion(String srcFile, String destFilePath, String dataDate) throws Exception {
		logger.info("-------conversion file Begining-------");
		logger.info("uploadAbsolutePath:" + srcFile + ";  destFilePath:" + destFilePath);

		try{
			if (!new File(destFilePath).exists()) {
				new File(destFilePath).mkdirs();
			}
			Workbook wb = null;
			FileOutputStream fileOut_ = null;
			String destFile = destFilePath + File.separator + "提供行领导" + dataDate + ".xls";
			logger.info("destFile: " + destFile);

			try (InputStream inp = new FileInputStream("D:\\excelTool\\template\\template_190319.xls")) {
				wb = WorkbookFactory.create(inp);
				fileOut_ = new FileOutputStream(destFile);
				wb.write(fileOut_);
			}finally {
				wb.close();
				fileOut_.close();
			}

			Map<String, String> map = getMapping2(dataDate);
			logger.info("get sheet mapping---> "+map.toString());

			Workbook srcWb = null;
			try (InputStream inp = new FileInputStream(srcFile)) {
				srcWb = WorkbookFactory.create(inp);
			}

			Workbook destWb = null;
			try (InputStream inp = new FileInputStream(destFile)) {
				destWb = WorkbookFactory.create(inp);
			}

			logger.info("foreach sheetMap Begining--->");
			for (String key : map.keySet()) {
				String val = map.get(key);
				if (!key.equals("目录") && val.equals("")) {
					continue;
				}

				logger.info("copyRowData Begining---> "+key);
				FileOutputStream fileOut = new FileOutputStream(destFile);
				Sheet destSheet = destWb.getSheet(key);
				if(key.equals("目录")){
					updateDirectoryPage(dataDate, destSheet);
				}else {
					copyRowData(destSheet, srcWb.getSheet(val));
				}

				destWb.write(fileOut);
				fileOut.flush();
				fileOut.close();
				logger.info("copyRowData Completed---> "+key);
			}
			logger.info("---foreach sheetMap Completed---");

			if (srcWb != null) {
				srcWb.close();
			}

			if (destWb != null) {
				destWb.close();
			}
			logger.info("-------conversion file Completed-------");
		}catch (Exception e){
			throw new Exception("conversion Exception: "+e);
		}



	}

	private void updateDirectoryPage(String dataDate, Sheet destSheet) throws ParseException {
		if(destSheet.getPhysicalNumberOfRows() > 2){
			Cell dateCell = destSheet.getRow(1).getCell(1);
			Date tmpDate = new SimpleDateFormat("yyyyMMdd").parse(dataDate);
			String showDate=new SimpleDateFormat("yyyy年MM月dd日").format(tmpDate);
			if(dateCell != null){
				dateCell.setCellValue(showDate);
			}
		}
	}

	public void copyRowData(Sheet destSheet, Sheet srcSheet) {
		if(srcSheet != null && destSheet != null){
			int firstRowNum = srcSheet.getFirstRowNum() + 1;
			int lastRowNum = srcSheet.getLastRowNum();

			for (int i = firstRowNum; i <= lastRowNum; i++) {
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
							destCell.setCellType(CellType.STRING);
							break;
						case NUMERIC:
							destCell.setCellValue(cell.getNumericCellValue());
							destCell.setCellType(CellType.NUMERIC);
							break;
						case FORMULA:
							destCell.setCellFormula(cell.getCellFormula());
							destCell.setCellType(CellType.FORMULA);
							break;
						case BOOLEAN:
							destCell.setCellValue(cell.getBooleanCellValue());
							destCell.setCellType(CellType.BOOLEAN);
							break;
						case ERROR:
							destCell.setCellValue(cell.getErrorCellValue());
							destCell.setCellType(CellType.ERROR);
							break;
						default:
							destCell.setCellValue(cell.getStringCellValue());
							destCell.setCellType(CellType.STRING);
							break;
					}
				}
			}

			String nullString = null;
			Cell nullCell;
			if(destSheet.getSheetName().equals("R0030")){
				nullCell = destSheet.getRow(2).getCell(7);
				if(nullCell != null){
					nullCell.setCellValue(nullString);
				}
			}else if(destSheet.getSheetName().equals("R0041")){
				nullCell = destSheet.getRow(2).getCell(5);
				if(nullCell != null){
					nullCell.setCellValue(nullString);
				}
				Cell nullCell2 = destSheet.getRow(2).getCell(6);
				if(nullCell2 != null){
					nullCell2.setCellValue(nullString);
				}
			}else if(destSheet.getSheetName().equals("R0061N")){
				nullCell = destSheet.getRow(2).getCell(8);
				if(nullCell != null){
					nullCell.setCellValue(nullString);
				}
			}else if(destSheet.getSheetName().equals("R0062N")){
				nullCell = destSheet.getRow(2).getCell(7);
				if(nullCell != null){
					nullCell.setCellValue(nullString);
				}
			}else if(destSheet.getSheetName().equals("R0061")){
				nullCell = destSheet.getRow(2).getCell(7);
				if(nullCell != null){
					nullCell.setCellValue(nullString);
				}
			}else if(destSheet.getSheetName().equals("R0062")){
				nullCell = destSheet.getRow(2).getCell(7);
				if(nullCell != null){
					nullCell.setCellValue(nullString);
				}
			}else if(destSheet.getSheetName().equals("R0008")){
				nullCell = destSheet.getRow(2).getCell(7);
				if(nullCell != null){
					nullCell.setCellValue(nullString);
				}
				Cell nullCell2 = destSheet.getRow(2).getCell(9);
				if(nullCell2 != null){
					nullCell2.setCellValue(nullString);
				}
			}else if(destSheet.getSheetName().equals("R0009")){
				nullCell = destSheet.getRow(2).getCell(8);
				if(nullCell != null){
					nullCell.setCellValue(nullString);
				}
			}else if(destSheet.getSheetName().equals("R0010")){
				nullCell = destSheet.getRow(2).getCell(8);
				if(nullCell != null){
					nullCell.setCellValue(nullString);
				}
			}else if(destSheet.getSheetName().equals("R0012")){
				nullCell = destSheet.getRow(2).getCell(9);
				if(nullCell != null){
					nullCell.setCellValue(nullString);
				}
			}else if(destSheet.getSheetName().equals("R0013")){
				nullCell = destSheet.getRow(2).getCell(9);
				if(nullCell != null){
					nullCell.setCellValue(nullString);
				}
			}else if(destSheet.getSheetName().equals("R0014")){
				nullCell = destSheet.getRow(2).getCell(7);
				if(nullCell != null){
					nullCell.setCellValue(nullString);
				}
			}else if(destSheet.getSheetName().equals("R0016")){
				nullCell = destSheet.getRow(2).getCell(10);
				if(nullCell != null){
					nullCell.setCellValue(nullString);
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
		map.put("R0030", "R0030-2019_境内_" + srcDate);
		map.put("R0041", "R0041_境内_" + srcDate);
		map.put("主动负债", "ZDFZ_FAST_境内_" + srcDate);
		map.put("R0040", "R0040_境内_" + srcDate);
		map.put("R0061N", "R0061N-2017_境内_" + srcDate);
		map.put("R0062N", "R0062N-2017_境内_" + srcDate);
		map.put("R0061", "R0061-2019_境内_" + srcDate);
		map.put("R0062", "R0062-2017_境内_" + srcDate);
		map.put("R0008", "R0008-2019_境内_" + srcDate);
		map.put("R0009", "R0009-2017_境内_" + srcDate);
		map.put("R0010", "R0010-2019_境内_" + srcDate);
		map.put("R0011", "R0011-2017_境内_" + srcDate);
		map.put("R0012", "R0012-2017_境内_" + srcDate);
		map.put("R0013", "R0013-2017_境内_" + srcDate);
		map.put("R0014", "R0014-2017_境内_" + srcDate);
		map.put("R0016", "R0016-2017_境内_" + srcDate);
		map.put("R0017", "R0017-2019_境内_" + srcDate);

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

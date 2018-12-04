package com.demo.exceldemomvn.service;

import com.demo.exceldemomvn.util.PoiUtil;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Service;
import org.springframework.util.ResourceUtils;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import static org.springframework.util.Assert.isTrue;

@Service
public class ExcelConversion implements FileConversion {

	private final Logger logger = LoggerFactory.getLogger(ExcelConversion.class);

	@Override
	public void conversion(String srcFile, String destinationFilePath) {
		//File srcFile_ = null;
		String absolutePath = "";
		File file = null;
		try {
			file = new File(ResourceUtils.getURL("classpath:").getPath());
		} catch (FileNotFoundException e) {
			throw new RuntimeException("获取根目录失败，无法创建上传目录！");
		}

		absolutePath = file.getAbsolutePath();
		File uploadFile = new File(absolutePath, srcFile);

		/*try {
			srcFile_ = ResourceUtils.getFile("classpath:" + srcFile);
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		}*/
		isTrue(uploadFile.exists(), "Src File not exits.");

		if (!new File(destinationFilePath).exists()) {
			new File(destinationFilePath).mkdirs();
		}

		Map<String, String> mapping = null;
		try {
			mapping = getMapping(uploadFile,
                    ResourceUtils.getFile("classpath:excel_template/template_1.xls"));
		} catch (Exception e) {
			e.printStackTrace();
		}

		mapping.forEach((key, val) -> {
			System.out.println(key + ":" + val);
		});

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

		map = new HashMap<>(templateSetList.size());

		for (String shortSheetName : templateSetList) {
			map.put(getSheetNameByShort(shortSheetName, templateSheetNameList),
					getSheetNameByShort(shortSheetName, srcSheetNameList));
		}

		return map;
	}

	private List<String> conversionSheetName(List<String> sheetNameList) {
		List<String> setList = new ArrayList<>(sheetNameList.size());
		sheetNameList.forEach(sheetName -> {
			if (sheetName.length() >= 5) {
				setList.add(sheetName.substring(0, 5));
			} else {
				logger.info("sheetName error." + sheetName);
			}
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

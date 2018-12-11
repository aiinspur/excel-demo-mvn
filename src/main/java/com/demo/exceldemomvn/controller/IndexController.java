package com.demo.exceldemomvn.controller;

import java.util.Date;

import org.apache.commons.lang3.time.DateFormatUtils;
import org.apache.log4j.Logger;
import org.jboss.logging.Param;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.multipart.MultipartFile;

import com.demo.exceldemomvn.service.FileConversion;
import com.demo.exceldemomvn.util.FileHandleUtil;

@Controller
public class IndexController {

	private static final Logger logger = Logger.getLogger(IndexController.class);

	@Autowired
	FileConversion fileConversion;

//	@Value("${destinationFile:/tmp/test/}")
//	private String destinationFile;
//
//	@Value("${srcFilePath:/tmp/test/upload/}")
//	private String srcFilePath;

	@ResponseBody
	@RequestMapping(value = "/upload")
	public String upload(@RequestParam("file") MultipartFile file, String srcFilePath, String destinationFile,@RequestParam("srcDate") String dataDate) {
		logger.info("srcFilePath:" + srcFilePath + ";destinationFile:" + destinationFile);

		String uploadPath = "";
		try {
			uploadPath = FileHandleUtil.upload2(file.getInputStream(), file.getOriginalFilename(), srcFilePath);

			fileConversion.conversion(uploadPath, destinationFile,dataDate);
		} catch (Exception e) {
			logger.info("upload Exception.");
			e.printStackTrace();
		}

		return "处理成功。报表文件存储路径为：" + destinationFile;

	}

	@ResponseBody
	@GetMapping("test")
	public String test(String srcFile, String destinationFile) throws Exception {
		//fileConversion.conversion(srcFile, destinationFile);
		return "hello test.";
	}

	@GetMapping
	public String index(Model model) {
		model.addAttribute("curDate", DateFormatUtils.format(new Date(), "yyyyMMdd"));
		return "index";
	}

}

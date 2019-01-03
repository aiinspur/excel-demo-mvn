package com.demo.exceldemomvn.controller;

import com.demo.exceldemomvn.service.FileConversion;
import com.demo.exceldemomvn.util.FileHandleUtil;
import org.apache.commons.lang3.time.DateFormatUtils;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.multipart.MultipartFile;

import java.util.Date;

@Controller
public class IndexController {
    private final Logger logger = LoggerFactory.getLogger(IndexController.class);

    @Autowired
    FileConversion fileConversion;

    @ResponseBody
    @RequestMapping(value = "/upload")
    public String upload(@RequestParam("file") MultipartFile file, String srcFilePath, String destFilePath, @RequestParam("srcDate") String dataDate) {
        logger.info("srcFilePath:" + srcFilePath + "; destFilePath:" + destFilePath);

        String uploadAbsolutePath = "";
        try {
            uploadAbsolutePath = FileHandleUtil.upload2(file.getInputStream(), file.getOriginalFilename(), srcFilePath);

            fileConversion.conversion(uploadAbsolutePath, destFilePath, dataDate);
        } catch (Exception e) {
            //e.printStackTrace();
            logger.error("processing excel Exception: " + e);
        }

        return "处理成功。报表文件存储路径为：" + destFilePath;

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

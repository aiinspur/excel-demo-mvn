package com.demo.exceldemomvn.controller;


import com.demo.exceldemomvn.service.FileConversion;
import com.demo.exceldemomvn.util.FileHandleUtil;
import org.apache.log4j.Logger;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Value;

import org.apache.log4j.Logger;

import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;


@RestController
public class FileController {
    @Autowired
    FileConversion fileConversion;

    private static final Logger logger = Logger.getLogger(FileController.class);

    @Value("${destinationFile:}")
    private String destinationFile;


    @RequestMapping(value = "/upload")
    public String upload(@RequestParam("file") MultipartFile file) {

        String uploadPath ="";
        try {
            uploadPath = FileHandleUtil.upload(file.getInputStream(), file.getOriginalFilename());

            fileConversion.conversion(uploadPath, destinationFile);
        } catch (Exception e) {
            e.printStackTrace();
        }

        return "文件存放路径为" + uploadPath;

    }

}

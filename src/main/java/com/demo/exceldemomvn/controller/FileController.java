package com.demo.exceldemomvn.controller;

import org.apache.log4j.Logger;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import com.demo.exceldemomvn.util.FileHandleUtil;

import java.io.IOException;

@RestController
public class FileController {

    private static final Logger logger = Logger.getLogger(FileController.class);

    @RequestMapping(value = "/upload")
    public String upload(@RequestParam("file") MultipartFile file) {

        String uploadPath ="";
        try {
            uploadPath = FileHandleUtil.upload(file.getInputStream(), file.getOriginalFilename());
        } catch (IOException e) {
            e.printStackTrace();
        }

        
        //doEx
        
        return "文件存放路径为" + uploadPath;

    }

}

package com.demo.exceldemomvn.util;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.util.Date;

import org.apache.commons.io.FileUtils;
import org.apache.commons.lang3.time.DateFormatUtils;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.util.ResourceUtils;

/**
 * SpringBoot上传文件工具类
 *
 * @author xipengfei
 */
public class FileHandleUtil {
    private static final Logger logger = LoggerFactory.getLogger(FileHandleUtil.class);
    /**
     * 绝对路径
     **/
    private static String absolutePath = "";

    /**
     * 文件存放的目录
     **/
    private static String fileDir = "upload/";

    public static String upload2(InputStream inputStream, String filename, String filepath) throws Exception {
        logger.info("-------upload file Beginning-----");
        String uploadAbsolutePath = "";
        try {
            File path = new File(filepath);
            if (!path.exists()) {
                path.mkdirs();
            }

            logger.info("orgFileName--->"+filename);
            uploadAbsolutePath = filepath + File.separator + transFileName(filename);
            logger.info("uploadAbsolutePath: " + uploadAbsolutePath);

            // 存文件
            File uploadFile = new File(uploadAbsolutePath);
            logger.info("uploadFile-------> " + uploadFile.getAbsoluteFile());

            FileUtils.copyInputStreamToFile(inputStream, uploadFile);
            logger.info("-------uploadFile Completed!!!-------");
        } catch (Exception e) {
            throw new Exception("upload2 Exception: " + e);
        }

        return uploadAbsolutePath;
    }

    /**
     * 上传单个文件 最后文件存放路径为：static/upload/image/test.jpg
     * 文件访问路径为：http://127.0.0.1:8080/upload/image/test.jpg
     * 该方法返回值为：/upload/image/test.jpg
     *
     * @param inputStream 文件流
     * @param filename    文件名，如：test.jpg
     * @return 成功：上传后的文件访问路径，失败返回：null
     */
    public static String upload(InputStream inputStream, String filename) {
        // 第一次会创建文件夹
        createDirIfNotExists();

        String resultPath = fileDir + transFileName(filename);

        // 存文件
        File uploadFile = new File(absolutePath, resultPath);
        try {
            FileUtils.copyInputStreamToFile(inputStream, uploadFile);
        } catch (IOException e) {
            e.printStackTrace();
            return null;
        }

        return resultPath;
    }

    private static String transFileName(String fileName) {
        return DateFormatUtils.format(new Date(), "yyyyMMddHHmmss") + "_" + fileName;
    }

    /**
     * 创建文件夹路径
     */
    private static void createDirIfNotExists() {
        if (!absolutePath.isEmpty()) {
            return;
        }

        // 获取跟目录
        File file = null;
        try {
            file = new File(ResourceUtils.getURL("classpath:").getPath());
        } catch (FileNotFoundException e) {
            throw new RuntimeException("获取根目录失败，无法创建上传目录！");
        }
        if (!file.exists()) {
            file = new File("");
        }

        absolutePath = file.getAbsolutePath();

        File upload = new File(absolutePath, fileDir);
        if (!upload.exists()) {
            upload.mkdirs();
        }
    }

    /**
     * 删除文件
     *
     * @param path 文件访问的路径upload开始 如： /upload/image/test.jpg
     * @return true 删除成功； false 删除失败
     */
    public static boolean delete(String path) {
        File file = new File(absolutePath, path);
        if (file.exists()) {
            return file.delete();
        }
        return false;
    }
}
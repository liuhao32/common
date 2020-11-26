package com.example.demo.annotation.controller;

import com.example.demo.annotation.Device;
import com.example.demo.annotation.ExcelExportUtil;
import com.example.demo.annotation.ExcelImportUtil;
import com.example.demo.annotation.ImportResult;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import java.io.File;
import java.io.FileOutputStream;
import java.io.InputStream;

/**
 * Copyright (C), 2020
 *
 * @author: liuhao
 * @date: 2020/11/21 22:57
 * @description:
 */
@RestController
public class TestController {


    @PostMapping("/test")
    public void test(MultipartFile file) throws Exception {

        InputStream inputStream = file.getInputStream();
        ImportResult<?> result = ExcelImportUtil.importExcel(inputStream, Device.class);
        System.out.println(result);

        String fileName = System.currentTimeMillis() + ".xls";
        File exportFile = new File(fileName);

        FileOutputStream output = new FileOutputStream(exportFile);
        ExcelExportUtil.exportExcel(output, result.getFails(), result.getFailPos(), result.getFailCause());
    }
}

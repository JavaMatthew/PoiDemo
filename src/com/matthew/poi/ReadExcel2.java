package com.matthew.poi;
/*
 * @(#)CreateExcel.java 2019年8月23日下午3:03:01
 * poiDemo
 * Copyright 2019 Thuisoft, Inc. All rights reserved.
 * THUNISOFT PROPRIETARY/CONFIDENTIAL. Use is subject to license terms.
 */


import java.io.FileInputStream;
import java.io.InputStream;

import org.apache.poi.hssf.extractor.ExcelExtractor;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;


/**
 * CreateExcel
 * @author Administrator
 * @version 1.0
 *读取excel的简便方法
 */
public class ReadExcel2 {

    /**
     * @param args
     * @throws Exception 
     */
    public static void main(String[] args) throws Exception {
        InputStream inputStream = new FileInputStream("E://用poi创建的工作簿.xls");
        POIFSFileSystem fSystem = new POIFSFileSystem(inputStream);
        HSSFWorkbook workbook = new HSSFWorkbook(fSystem);
        
        ExcelExtractor excelExtractor = new ExcelExtractor(workbook);
        excelExtractor.setIncludeSheetNames(false);//是否输出sheet页名字
        excelExtractor.setIncludeBlankCells(true);//是否输出空白单元格
        System.out.println(excelExtractor.getText());
        excelExtractor.close();
        
    }
}

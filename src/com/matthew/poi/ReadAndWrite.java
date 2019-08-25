/*
 * @(#)CreateCell.java 2019年8月23日下午3:25:36
 * poiDemo
 * Copyright 2019 Thuisoft, Inc. All rights reserved.
 * THUNISOFT PROPRIETARY/CONFIDENTIAL. Use is subject to license terms.
 */
package com.matthew.poi;


import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * CreateCell
 * @author Administrator
 * @version 1.0
 *设置边框
 */
public class ReadAndWrite {

    /**
     * @param args
     */
    public static void main(String[] args) throws Exception{
       InputStream inputStream = new FileInputStream("E://用poi创建的工作簿.xls");
       POIFSFileSystem fileSystem = new POIFSFileSystem(inputStream);
       Workbook workbook = new HSSFWorkbook(fileSystem);
       Sheet sheet = workbook.getSheetAt(0);
       Row row = sheet.getRow(0);//获取第一行
       Cell cell = row.getCell(0);//获取单元格
       
       
        
        FileOutputStream fileOutputStream = new FileOutputStream("E://字体样式.xls");
        workbook.write(fileOutputStream);
        workbook.close();
    }

}

/*
 * @(#)CreateCell.java 2019年8月23日下午3:25:36
 * poiDemo
 * Copyright 2019 Thuisoft, Inc. All rights reserved.
 * THUNISOFT PROPRIETARY/CONFIDENTIAL. Use is subject to license terms.
 */
package com.matthew.poi;

import java.io.FileOutputStream;
import java.util.Calendar;
import java.util.Date;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * CreateCell
 * @author Administrator
 * @version 1.0
 *
 */
public class CreateCell {

    /**
     * @param args
     */
    public static void main(String[] args) throws Exception{
        Workbook workbook = new HSSFWorkbook();//创建workbook对象，定义一个新的工作簿
        Sheet sheet1 = workbook.createSheet("第一个sheet页");//创建sheet页
        Sheet sheet2 = workbook.createSheet("第二个sheet页");
        Row row1 = sheet1.createRow(0);//创建行
        Row row2 = sheet2.createRow(2);
        Cell cell1 = row1.createCell(1);//创建单元格
        Cell cell2 = row2.createCell(2);
        cell1.setCellValue(new Date());//设置值
        cell2.setCellValue(Calendar.getInstance());
        
        //时间需要设置样式，否则显示的是一串数字
        CreationHelper creationHelper = workbook.getCreationHelper();
        CellStyle cellStyle = workbook.createCellStyle();//单元格样式
        cellStyle.setDataFormat(creationHelper.createDataFormat().getFormat("yyyy-mm-dd hh:mm:ss"));
        cell1.setCellStyle(cellStyle);
        cell2.setCellStyle(cellStyle);
        
        FileOutputStream fileOutputStream = new FileOutputStream("E://用poi创建的工作簿.xls");
        workbook.write(fileOutputStream);
        workbook.close();
    }

}

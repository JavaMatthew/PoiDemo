/*
 * @(#)CreateCell.java 2019年8月23日下午3:25:36
 * poiDemo
 * Copyright 2019 Thuisoft, Inc. All rights reserved.
 * THUNISOFT PROPRIETARY/CONFIDENTIAL. Use is subject to license terms.
 */
package com.matthew.poi;


import java.io.FileOutputStream;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * CreateCell
 * @author Administrator
 * @version 1.0
 *设置边框
 */
public class SetWord {

    /**
     * @param args
     */
    public static void main(String[] args) throws Exception{
        Workbook workbook = new HSSFWorkbook();//创建workbook对象，定义一个新的工作簿
        Sheet sheet1 = workbook.createSheet("第一个sheet页");//创建sheet页
        Row row1 = sheet1.createRow(3);//创建行
        
//        创建一个字体处理类
        Font font = workbook.createFont();
        font.setFontHeightInPoints((short)24);//字体大小
        font.setFontName("Courier New");//字体名字
        font.setItalic(true);//斜体
        font.setStrikeout(true);//删除线
        
        CellStyle style = workbook.createCellStyle();
        style.setFont(font);
        
        
        Cell cell = row1.createCell((short)1);
        cell.setCellValue("This is test of fonts");
        cell.setCellStyle(style);
        
        FileOutputStream fileOutputStream = new FileOutputStream("E://字体样式.xls");
        workbook.write(fileOutputStream);
        workbook.close();
    }

}

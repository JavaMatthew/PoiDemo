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
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * CreateCell
 * @author Administrator
 * @version 1.0
 *设置背景色
 */
public class SetBackground {

    /**
     * @param args
     */
    public static void main(String[] args) throws Exception{
        Workbook workbook = new HSSFWorkbook();//创建workbook对象，定义一个新的工作簿
        Sheet sheet1 = workbook.createSheet("第一个sheet页");//创建sheet页
        Row row1 = sheet1.createRow(3);//创建行
        
        Cell cell = row1.createCell(3);
        cell.setCellValue("XX");
        CellStyle cellStyle = workbook.createCellStyle();
        //注意：setFillBackgroundColor没有效果，在网上也没有相应的解释
        cellStyle.setFillBackgroundColor(IndexedColors.WHITE.getIndex());
        cellStyle.setFillPattern(CellStyle.ALIGN_CENTER);   //注意这个参数如果不设置，背景色将不会显示
        
        Cell cell2 = row1.createCell(4);
        cell2.setCellValue("XX");
        CellStyle cellStyle2 = workbook.createCellStyle();
        cellStyle2.setFillForegroundColor(IndexedColors.DARK_RED.getIndex());//前景色
        cellStyle2.setFillPattern(CellStyle.SOLID_FOREGROUND);
        
        
        cell2.setCellStyle(cellStyle2);
        cell.setCellStyle(cellStyle);
        FileOutputStream fileOutputStream = new FileOutputStream("E://对齐方式.xls");
        workbook.write(fileOutputStream);
        workbook.close();
    }

}

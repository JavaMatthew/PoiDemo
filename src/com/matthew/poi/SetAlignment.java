/*
 * @(#)CreateCell.java 2019年8月23日下午3:25:36
 * poiDemo
 * Copyright 2019 Thuisoft, Inc. All rights reserved.
 * THUNISOFT PROPRIETARY/CONFIDENTIAL. Use is subject to license terms.
 */
package com.matthew.poi;

import java.io.FileOutputStream;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * CreateCell
 * @author Administrator
 * @version 1.0
 *设置对其方式
 */
public class SetAlignment {

    /**
     * @param args
     */
    public static void main(String[] args) throws Exception{
        Workbook workbook = new HSSFWorkbook();//创建workbook对象，定义一个新的工作簿
        Sheet sheet1 = workbook.createSheet("第一个sheet页");//创建sheet页
        Row row1 = sheet1.createRow(0);//创建行
        
        sheet1.setColumnWidth(0, 256*30+184);//sheet.setColumnWidth(0, 256*width+184);
        sheet1.setColumnWidth(1, 256*30+184);//sheet.setColumnWidth(0, 256*width+184);
        sheet1.setColumnWidth(2, 256*30+184);//sheet.setColumnWidth(0, 256*width+184);
        sheet1.setColumnWidth(3, 256*30+184);//sheet.setColumnWidth(0, 256*width+184);
        row1.setHeightInPoints(30);
        createCell(workbook, row1, (short)0, HSSFCellStyle.ALIGN_CENTER, HSSFCellStyle.VERTICAL_BOTTOM);
        createCell(workbook, row1, (short)1, HSSFCellStyle.ALIGN_JUSTIFY, HSSFCellStyle.VERTICAL_JUSTIFY);
        createCell(workbook, row1, (short)2, HSSFCellStyle.ALIGN_FILL, HSSFCellStyle.VERTICAL_CENTER);
        createCell(workbook, row1, (short)3, HSSFCellStyle.ALIGN_GENERAL, HSSFCellStyle.VERTICAL_TOP);
        
        FileOutputStream fileOutputStream = new FileOutputStream("E://对齐方式.xls");
        workbook.write(fileOutputStream);
        workbook.close();
    }
    
    /**
     * 创建一个单元格并为其设定指定的对其方式
     * @param workbook  工作簿
     * @param row   行
     * @param column    列
     * @param halign    水平方向对其方式
     * @param valign    垂直方向对其方式
     */
    private static void createCell(Workbook workbook,Row row,short column,short halign,short valign) {
        Cell cell = row.createCell(column);//创建单元格
        cell.setCellValue("Align it");//设置值
        CellStyle cellStyle = workbook.createCellStyle();//创建单元格样式
        cellStyle.setAlignment(halign);//设置单元格水平方向对其方式
        cellStyle.setVerticalAlignment(valign);//设置单元格垂直方向对其方式
        cell.setCellStyle(cellStyle);//设置单元格样式
        
    }

}

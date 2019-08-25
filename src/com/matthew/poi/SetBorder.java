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
 *设置边框
 */
public class SetBorder {

    /**
     * @param args
     */
    public static void main(String[] args) throws Exception{
        Workbook workbook = new HSSFWorkbook();//创建workbook对象，定义一个新的工作簿
        Sheet sheet1 = workbook.createSheet("第一个sheet页");//创建sheet页
        Row row1 = sheet1.createRow(3);//创建行
        
        Cell cell = row1.createCell(1);
        cell.setCellValue(4);
        CellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setBorderBottom(CellStyle.BORDER_THIN);//底部边框
        cellStyle.setBottomBorderColor(IndexedColors.BLACK.getIndex());//底部颜色
        
        cellStyle.setBorderLeft(CellStyle.BORDER_THIN);//左边框
        cellStyle.setLeftBorderColor(IndexedColors.RED.getIndex());//左边框颜色
        
        cellStyle.setBorderRight(CellStyle.BORDER_THIN);//右边框
        cellStyle.setRightBorderColor(IndexedColors.GREEN.getIndex());//右边框颜色
        
        cellStyle.setBorderTop(CellStyle.BORDER_MEDIUM_DASHED);//上边框
        cellStyle.setTopBorderColor(IndexedColors.BLUE.getIndex());//上边框颜色 
        
        cell.setCellStyle(cellStyle);
        FileOutputStream fileOutputStream = new FileOutputStream("E://对齐方式.xls");
        workbook.write(fileOutputStream);
        workbook.close();
    }

}

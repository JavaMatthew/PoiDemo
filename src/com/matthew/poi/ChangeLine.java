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
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * CreateCell
 * @author Administrator
 * @version 1.0
 *单元格换行
 */
public class ChangeLine {

    /**
     * @param args
     */
    public static void main(String[] args) throws Exception{
	    Workbook workbook = new HSSFWorkbook();//创建workbook对象，定义一个新的工作簿
        Sheet sheet1 = workbook.createSheet("第一个sheet页");//创建sheet页
        Row row1 = sheet1.createRow(3);//创建行
        Cell cell = row1.createCell(1);
        cell.setCellValue("我要换行成功了么");
        
        CellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setWrapText(true);//设置可以换行
        cell.setCellStyle(cellStyle);
        
        //调整一下行高
        row1.setHeightInPoints(2 * sheet1.getDefaultRowHeightInPoints());
        //调整单元格宽度
        sheet1.autoSizeColumn(2);
        
		FileOutputStream fileOutputStream = new FileOutputStream("E://单元格换行.xls");
		workbook.write(fileOutputStream);
		workbook.close();
    }

}

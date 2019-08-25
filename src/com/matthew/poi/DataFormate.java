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
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * CreateCell
 * @author Administrator
 * @version 1.0
 *数据格式
 */
public class DataFormate {

    /**
     * @param args
     */
    public static void main(String[] args) throws Exception{
	    Workbook workbook = new HSSFWorkbook();//创建workbook对象，定义一个新的工作簿
        Sheet sheet1 = workbook.createSheet("第一个sheet页");//创建sheet页

        Row row;
        Cell cell;
        CellStyle cellStyle;
        DataFormat format = workbook.createDataFormat();
        short rowNum = 0;
        short colNum = 0;
        
        row=sheet1.createRow(rowNum++);
        cell=row.createCell(colNum);
        cell.setCellValue(11112321.32);
        cellStyle = workbook.createCellStyle();
        cellStyle.setDataFormat(format.getFormat("0.0"));//设置数据格式
        cell.setCellStyle(cellStyle);
        
        row=sheet1.createRow(rowNum++);
        cell=row.createCell(colNum);
        cell.setCellValue(111112321.32);
        cellStyle = workbook.createCellStyle();
        cellStyle.setDataFormat(format.getFormat("#,##0.000"));
        cell.setCellStyle(cellStyle);
        
        
		FileOutputStream fileOutputStream = new FileOutputStream("E://数据格式.xls");
		workbook.write(fileOutputStream);
		workbook.close();
    }

}

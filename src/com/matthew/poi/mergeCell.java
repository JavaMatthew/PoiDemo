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
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;

/**
 * CreateCell
 * @author Administrator
 * @version 1.0
 *合并单元格
 */
public class mergeCell {

    /**
     * @param args
     */
    public static void main(String[] args) throws Exception{
        Workbook workbook = new HSSFWorkbook();//创建workbook对象，定义一个新的工作簿
        Sheet sheet1 = workbook.createSheet("第一个sheet页");//创建sheet页
        Row row1 = sheet1.createRow(0);//创建行
        
//        Cell cell2 = row1.createCell(2);
//        cell2.setCellValue("XX");
        
        Cell cell = row1.createCell(1);
        cell.setCellValue("单元格合并测试");
        //注意：合并单元格后，只会显示左上角单元格的值
        //例如：下面的合并是B1，B2，C1，C2，合并后的单元格是B1
        sheet1.addMergedRegion(new CellRangeAddress(
            0,      //起始行
            1,      //结束行
            1,      //起始列
            2));    //结束列
        
        
        
        FileOutputStream fileOutputStream = new FileOutputStream("E://合并单元格.xls");
        workbook.write(fileOutputStream);
        workbook.close();
    }

}

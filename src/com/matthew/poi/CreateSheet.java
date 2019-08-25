/*
 * @(#)CreateSheet.java 2019年8月23日下午3:09:00
 * poiDemo
 * Copyright 2019 Thuisoft, Inc. All rights reserved.
 * THUNISOFT PROPRIETARY/CONFIDENTIAL. Use is subject to license terms.
 */
package com.matthew.poi;

import java.io.FileOutputStream;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * CreateSheet
 * @author Administrator
 * @version 1.0
 *
 */
public class CreateSheet {

    /**
     * @param args
     * @throws Exception 
     */
    public static void main(String[] args) throws Exception {
        Workbook workbook = new HSSFWorkbook();
        workbook.createSheet("我的第一个SHEET页");
        workbook.createSheet("我的第二个SHEET页");
        FileOutputStream fileOutputStream = new FileOutputStream("E://用poi创建的工作簿.xls");
        workbook.write(fileOutputStream);
        workbook.close();
    }

}

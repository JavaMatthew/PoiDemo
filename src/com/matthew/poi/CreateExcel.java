package com.matthew.poi;
/*
 * @(#)CreateExcel.java 2019年8月23日下午3:03:01
 * poiDemo
 * Copyright 2019 Thuisoft, Inc. All rights reserved.
 * THUNISOFT PROPRIETARY/CONFIDENTIAL. Use is subject to license terms.
 */


import java.io.FileOutputStream;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * CreateExcel
 * @author Administrator
 * @version 1.0
 *
 */
public class CreateExcel {

    /**
     * @param args
     * @throws Exception 
     */
    public static void main(String[] args) throws Exception {
        Workbook workbook = new HSSFWorkbook();
        FileOutputStream fileOutputStream = new FileOutputStream("E://用poi创建的工作簿.xls");
        workbook.write(fileOutputStream);
        workbook.close();
    }

}

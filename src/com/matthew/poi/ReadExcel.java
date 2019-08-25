package com.matthew.poi;
/*
 * @(#)CreateExcel.java 2019年8月23日下午3:03:01
 * poiDemo
 * Copyright 2019 Thuisoft, Inc. All rights reserved.
 * THUNISOFT PROPRIETARY/CONFIDENTIAL. Use is subject to license terms.
 */


import java.io.FileInputStream;
import java.io.InputStream;
import java.text.SimpleDateFormat;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.DateUtil;


/**
 * CreateExcel
 * @author Administrator
 * @version 1.0
 *读取Excel
 */
public class ReadExcel {

    /**
     * @param args
     * @throws Exception 
     */
    public static void main(String[] args) throws Exception {
        InputStream inputStream = new FileInputStream("E://用poi创建的工作簿.xls");
        POIFSFileSystem fSystem = new POIFSFileSystem(inputStream);
        HSSFWorkbook workbook = new HSSFWorkbook(fSystem);
        
//        HSSFSheet hssfSheet = workbook.getSheetAt(0);//获取第一个Sheet页
        if (workbook.getNumberOfSheets() == 0) {
            workbook.close();
            return;
        }
        for(int sheetNum=0; sheetNum < workbook.getNumberOfSheets(); sheetNum++) {
            HSSFSheet hssfSheet = workbook.getSheetAt(sheetNum);
            System.out.println(hssfSheet.getSheetName());
            //遍历行row
            for(int rowNum = 0; rowNum <= hssfSheet.getLastRowNum(); rowNum++) {
                HSSFRow hssfRow = hssfSheet.getRow(rowNum);
                if (hssfRow == null) {
                    continue;
                }
                System.out.print("这是第"+ (rowNum+1) +"行： ");
                //遍历列Cell
                for(int CellNum=0; CellNum <= hssfRow.getLastCellNum(); CellNum++) {
                    HSSFCell hssfCell = hssfRow.getCell(CellNum);
                    if (hssfCell == null) {
                        continue;
                    }
                    System.out.print("  " + getValue(hssfCell));
                }
                System.out.println();
            }
            System.out.println();
        }
        workbook.close();
    }

    
    private static String getValue(HSSFCell hssfCell) {
        switch(hssfCell.getCellType()) {
            case HSSFCell.CELL_TYPE_BOOLEAN:
                return String.valueOf(hssfCell.getBooleanCellValue());
            case HSSFCell.CELL_TYPE_NUMERIC:
                if (DateUtil.isCellDateFormatted(hssfCell)) {
                    SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy-mm-dd hh:mm:ss");
                    return simpleDateFormat.format(hssfCell.getDateCellValue());
                }
                return String.valueOf(hssfCell.getNumericCellValue());
            case HSSFCell.CELL_TYPE_FORMULA:
                return String.valueOf(hssfCell.getBooleanCellValue());
            default:
                return hssfCell.getStringCellValue();
        }
    }
}

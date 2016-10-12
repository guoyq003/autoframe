package com.guoyq.auto.util;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.LinkedHashMap;

public class DataUtilDemo {

    int excelRowNum;//excel行号
    int excelColumnNum;//excel列号

    private static DataUtilDemo instance=null;
    //单例模式
    public static DataUtilDemo getInstance(){
        if (instance==null){
            instance=new DataUtilDemo();
        }
        return instance;
    }
    //获取测试数据文件夹的绝对路径
    public String getTestDateFolderPath(String dataFolderPath){
        String testDataFilePath=null;
        File directory=new File(dataFolderPath);
        try {
            testDataFilePath=directory.getAbsolutePath();
        }catch (Exception e){
            e.printStackTrace();
        }
        return testDataFilePath;
    }
    //获取测试数据文件路径
    public String getTestDateFilePath(String dataFolderPath,String fileName){
        String testDateFilePath=this.getTestDateFolderPath(dataFolderPath)+File.separator+fileName;
        return testDateFilePath;
    }
    //获取dataExcel文件
    public Workbook getWorkbook(String testDataFilePath){
        Workbook wb = null;
        boolean isE2007 = false;
        if (testDataFilePath.endsWith("xlsx")) {
            isE2007 = true;
        }
        InputStream input = null;
        try {

            input = new FileInputStream(testDataFilePath);
            //文件格式(2003或者2007)来初始化
            if (isE2007) {
                wb = new XSSFWorkbook(input);
            } else {
                wb = new HSSFWorkbook(input);
            }
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            try {
                input.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
        return wb;
    }
    //获取dataExcel文件
    public Workbook getHSSFWorkbook(String testDataFilePath) {
        Workbook wb = this.getWorkbook(testDataFilePath);
        return wb;
    }
    //获取EXCEl表格中的指定sheet
    public Sheet getSheet(String testDataFilePath,int sheetNum){
        Workbook wb = this.getHSSFWorkbook(testDataFilePath);
        Sheet sheet = wb.getSheetAt(sheetNum);
        return sheet;
    }
    //根据所在的列名获取所在的列号
//    public int getColumnCount(String columnName,String testDataFilePath,int sheetNum){
//        Sheet sheet = this.getSheet(testDataFilePath,sheetNum);
//        for (int i=0;i<1;i++){
//        }
//    }
    //初始化excel中的数字double为整数
    //在TestData.xls文件中，根据columnName，获取所在的列号
    public int getColumnCount(String columnName,String excelFileName,int sheetNum) {
        Sheet sheet = this.getSheet(excelFileName,sheetNum);
        for (int i = 0; i < sheet.getPhysicalNumberOfRows(); i++) {
            Row row = sheet.getRow(i);
            if (row != null) {
                sheet.getColumnBreaks();
                for (int j = 0; j < row.getPhysicalNumberOfCells(); j++) {
                    Cell cell = row.getCell(j);
                    if (cell != null) {
                        if (cell.toString().equals(columnName)) {
                            this.excelColumnNum = j;
                            break;
                        }
                    }
                }
            }
        }
        return this.excelColumnNum;
    }
    //在TestData.xls文件中，根据rowname（场景名），获取所在的行号
    public int getRowCount(String rowName,String excelFileName,int sheetNum) {
        Sheet sheet = this.getSheet(excelFileName,sheetNum);
        for (int i = 1; i < sheet.getPhysicalNumberOfRows(); i++) {
            Row row = sheet.getRow(i);
            if (row != null) {
                sheet.getColumnBreaks();
                for (int j = 0; j < row.getPhysicalNumberOfCells(); j++) {
                    Cell cell = row.getCell(j);
                    if (cell != null) {
                        if (cell.toString().equals(rowName)) {
                            this.excelRowNum = i;
                            break;
                        }
                    }
                }
            }
        }
        return this.excelRowNum;
    }
    //在TestData.xls中，根据rowName（场景名称）和列名columName，获得对应单元格的值
    public String getCellValue(String rowName, String columnName,String excelFileName,int sheetNum) {
        String cellValue = null;
        excelRowNum = this.getRowCount(rowName,excelFileName,sheetNum);
        excelColumnNum = this.getColumnCount(columnName,excelFileName,sheetNum);
        Sheet sheet = this.getSheet(excelFileName,sheetNum);
        Row row = sheet.getRow(excelRowNum);
        if (row != null) {
            sheet.getColumnBreaks();
            Cell cell = row.getCell(excelColumnNum);
            if (cell == null)
                cellValue = "";
            else
                cellValue = cell.toString();
        }
        return cellValue;
    }
}
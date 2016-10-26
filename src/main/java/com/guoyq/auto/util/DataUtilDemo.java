package com.guoyq.auto.util;


import org.apache.log4j.Logger;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.*;

public class DataUtilDemo {

    public static Logger logger = Logger.getLogger(DataUtilDemo.class);
    int excelRowNum;//excel行号
    int excelColumnNum;//excel列号
//    //获取测试场景
//    private List<Map<String, String>> senarioList = new ArrayList<Map<String, String>>();
//
//    public List<Map<String, String>> getSenarioList() {
//        return senarioList;
//    }
//
//    public void setSenarioList(List<Map<String, String>> senarioList) {
//        this.senarioList = senarioList;
//    }

    private static DataUtilDemo instance=null;
    //单例模式
    public static DataUtilDemo getInstance(){
        if (instance==null){
            instance=new DataUtilDemo();
        }
        return instance;
    }
    //获取测试数据文件夹的绝对路径
//    public String getTestDateFolderPath(String dataFolderPath){
//        String testDataFilePath=null;
//        File directory=new File(dataFolderPath);
//        try {
//            testDataFilePath=directory.getAbsolutePath();
//        }catch (Exception e){
//            e.printStackTrace();
//        }
//        return testDataFilePath;
//    }
    //获取测试数据文件路径
//    public String getTestDateFilePath(String dataFolderPath,String fileName){
//        String testDateFilePath=this.getTestDateFolderPath(dataFolderPath)+File.separator+fileName;
//        return testDateFilePath;
//    }
    //获取dataExcel文件
    public Workbook getWorkbook(String excelFilePath){
        //excelFilePath 读取的excel文件路径
        Workbook wb = null;
        boolean isE2007 = false;
        if (excelFilePath.endsWith("xlsx")) {
            isE2007 = true;
        }
        InputStream input = null;
        try {

            input = new FileInputStream(excelFilePath);
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
    public Workbook getHSSFWorkbook(String excelFilePath) {
        Workbook wb = this.getWorkbook(excelFilePath);
        return wb;
    }
    //获取EXCEl表格中的指定sheet
    public Sheet getSheet(String excelFilePath,int sheetNum){
        Workbook wb = this.getHSSFWorkbook(excelFilePath);
        Sheet sheet = wb.getSheetAt(sheetNum);
        return sheet;
    }

    //在TestData.xls文件中，根据columnName，获取所在的列号
    public int getColumnCount(String columnName,String excelFilePath,int sheetNum) {
        Sheet sheet = this.getSheet(excelFilePath,sheetNum);
        for (int i = 0; i < sheet.getPhysicalNumberOfRows(); i++) {
            Row row = sheet.getRow(i);
            if (row != null) {
                sheet.getColumnBreaks();
                List<String> cellList=new ArrayList<String>();
                for (int j=0; j < row.getPhysicalNumberOfCells(); j++) {
                    Cell cell = row.getCell(j);
                    cellList.add(new FormatExcelCell().formatExcelCelltoString(cell));
                }
                if (cellList.contains(columnName)){
                    excelColumnNum=cellList.indexOf(columnName);
                    break;
                }else {
                    //-1表示没有查到对应的数据
                    excelColumnNum=-1;
                }
            }
        }
        return this.excelColumnNum;
    }
    //在TestData.xls文件中，根据rowname（场景名），获取所在的行号
    public int getRowCount(String rowName,String excelFilePath,int sheetNum) {
        Sheet sheet = this.getSheet(excelFilePath,sheetNum);
        for (int i = 0; i < sheet.getPhysicalNumberOfRows(); i++) {
            Row row = sheet.getRow(i);
            if (row != null) {
                sheet.getColumnBreaks();
                List<String> cellList=new ArrayList<String>();
                for (int j = 0; j < row.getPhysicalNumberOfCells(); j++) {
                    Cell cell = row.getCell(j);
                    cellList.add(new FormatExcelCell().formatExcelCelltoString(cell));
                }
                if (cellList.contains(rowName)){
                    excelRowNum=row.getRowNum();
                    break;
                }else {
                    //-1表示没有查到对应的数据
                    excelRowNum=-1;
                }
            }
        }
        return this.excelRowNum;
    }
    //在TestData.xls中，根据rowName（场景名称）和列名columName，获得对应单元格的值
    public String getCellValue(String rowName, String columnName,String excelFilePath,int sheetNum) {
        String cellValue = null;
        excelRowNum = this.getRowCount(rowName,excelFilePath,sheetNum);
        excelColumnNum = this.getColumnCount(columnName,excelFilePath,sheetNum);
        Sheet sheet = this.getSheet(excelFilePath,sheetNum);
        Row row = sheet.getRow(excelRowNum);
        if (row != null) {
            sheet.getColumnBreaks();
            Cell cell = row.getCell(excelColumnNum);
            if (cell == null)
                cellValue = "";
            else
                cellValue = new FormatExcelCell().formatExcelCelltoString(cell);
        }
        return cellValue;
    }
    /**
     * 在TestData.xls中，设置rowName（场景名称）和列名columnName对应的单元格的值
     */
//    public void setCellValue(String cellValue,String rowName, String columnName,String excelFilePath,int sheetNum) {
//        Workbook wb = this.getHSSFWorkbook(excelFilePath);
//        try {
//            Sheet sheet = wb.getSheetAt(0);
//            Row row = sheet.getRow(getRowCount(rowName,excelFilePath,sheetNum));
//            System.out.println(row);
//            if (row != null) {
//                Cell cell = row.getCell(getColumnCount(columnName,excelFilePath,sheetNum));
//                System.out.println(cell);
//                if (cell == null) {
//                    cell = row.createCell(getColumnCount(columnName,excelFilePath,sheetNum));
//                }
//                cell.setCellValue(cellValue);
//            }
//            FileOutputStream fos = new FileOutputStream(excelFilePath);
//            fos.flush();
//            wb.write(fos);
//            fos.close();
//        } catch (IOException e) {
//            e.printStackTrace();
//        }
//    }
    //获取TestData.xls中sheet的前两列，即所有场景名称及类型
    public List<Map<String, String>> getSeniorName(String excelFilePath,int sheetNum) {
        List<Map<String, String>> list = new ArrayList<Map<String, String>>();
        Sheet sheet = this.getSheet(excelFilePath,sheetNum);
        try {
            for (int i = 1; i < sheet.getPhysicalNumberOfRows(); i++) {
                Row row = sheet.getRow(i);
                if (row != null) {
                    sheet.getColumnBreaks();
                    if (row.getCell(0).toString() != "" && !"".equals(row.getCell(0).toString())) {
                        Map<String, String> seniorMap = new HashMap<String, String>();
                        String seniorName = new FormatExcelCell().formatExcelCelltoString(row.getCell(0));
                        String seniorProtocol = new FormatExcelCell().formatExcelCelltoString(row.getCell(1));
                        seniorMap.put(seniorName, seniorProtocol);
                        list.add(seniorMap);
                    }
                }
            }
        } catch (NullPointerException e) {
            new Throwable();
        }
//        this.setSenarioList(list);
        return list;
    }
    //以targetFileName为名的文件，将文件的绝对路径add到filepath中
    public String findFiles(String baseDirName, String targetFileName, List filesPath) {

        File baseDir = new File(baseDirName); // 创建一个File对象
        if (!baseDir.exists() || !baseDir.isDirectory()) {
            logger.error("文件查找失败：" + baseDirName + "不是一个目录！");
        }
        String tempName = null; // 判断目录是否存在
        File tempFile;
        File[] files = baseDir.listFiles();
        for (int i = 0; i < files.length; i++) {
            tempFile = files[i];
            if (tempFile.isDirectory()) {
                findFiles(tempFile.getAbsolutePath(), targetFileName, filesPath);
            } else if (tempFile.isFile()) {
                tempName = tempFile.getName();
                if (tempName.equalsIgnoreCase(targetFileName)) {
                    filesPath.add(tempFile.getAbsolutePath());
                }
            }
        }
        return tempName;
    }

}
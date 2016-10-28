package com.guoyq.auto.util;


import com.guoyq.auto.config.Constants;
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
    private HashMap<String, LinkedHashMap<String, String>> testCaseMap = null;
    public static Logger logger = Logger.getLogger(DataUtilDemo.class);
    int excelRowNum;//excel行号
    int excelColumnNum;//excel列号

    private String senario = "";
    private String currentTestCase = "";
    public HashMap<String, LinkedHashMap<String, String>> getTestCaseMap() {
        return testCaseMap;
    }

    public void setTestCaseMap(HashMap<String, LinkedHashMap<String, String>> testCaseMap) {
        this.testCaseMap = testCaseMap;
    }

    private static DataUtilDemo instance=null;
    //单例模式
    public static DataUtilDemo getInstance(){
        if (instance==null){
            instance=new DataUtilDemo();
        }
        return instance;
    }
    //获取通过判断文件后缀确定是2007还是2003返回Workbook
    public Workbook getWorkbook(String excelFilePath){
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
        } catch (Exception e) {
            e.printStackTrace();
        }
        finally {
            try {
                input.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
        return wb;
    }
    //获取EXCEl表格中的指定sheet
    public Sheet getSheet(String excelFilePath,int sheetNum){
        Sheet sheet = this.getWorkbook(excelFilePath).getSheetAt(sheetNum);
        return sheet;
    }
    //获取场景名称
    public String getSenario(String excelFilePath,int sheetNum) {
        if (senario.equals(""))//默认取第一个
        {
            Map<String, String> stringStringMap = this.getSeniorName(excelFilePath,sheetNum).get(0);
            senario = stringStringMap.keySet().iterator().next();
        }
        return senario;
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
        return list;
    }
    /**
     * 根据filename，获取场景名对应的测试用例
     * 首先查找data目录下是否存在以filename为名的excel文件，如果不存在，则查找RunList.xls
     * 将excel中需要运行（RunFlag不为空且不为‘N’）的用例封装到LinkedHashMap<String, String>中，key为第二列TestCase，value为第一列+','+第四列描述
     *
     * @return LinkedHashMap<String, String>
     */
    public LinkedHashMap<String, String> getTestCase(String caseFilePath) {
        LinkedHashMap<String, String> testCaseMap = new LinkedHashMap<String, String>();
        try {
            Workbook wb = this.getWorkbook(caseFilePath);
            Sheet sheet = wb.getSheetAt(0);
            for (int i = 1; i < sheet.getPhysicalNumberOfRows(); i++) {
                Row row = sheet.getRow(i);
                if (row != null) {
                    sheet.getColumnBreaks();
                    if (row.getCell(2) == null || "N".equals(row.getCell(2).toString()))
                        continue;
                    if (row.getCell(1) == null || "".equals(row.getCell(1).toString()))
                        break;
                    String testCase = row.getCell(1).toString().trim();
                    String testDesc = "";
                    try {
                        testDesc = new FormatExcelCell().formatExcelCelltoString(row.getCell(0)).trim() + Constants.Comma + row.getCell(3).toString().trim();
                    } catch (Exception e) {
                        logger.error("获取运行列表值为空，"+e);
                    }
                    if (testCaseMap.containsKey(testCase)) {
                        testCase = testCase + Constants.DUP_MARK + i;
                    }
                    testCaseMap.put(testCase, testDesc);
                }
            }
        } catch (Exception e) {
            logger.error(e);
            e.printStackTrace();
        }
        return testCaseMap;
    }
}
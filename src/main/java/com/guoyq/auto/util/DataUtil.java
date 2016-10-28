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


public class DataUtil {
    public static Logger logger = Logger.getLogger(DataUtil.class);
    static DataUtil instance = null;
    private String senario = "";
    private String currentTestCase = "";

    private HashMap<String, LinkedHashMap<String, String>> testCaseMap = null;

    private List<Map<String, String>> senarioList = new ArrayList<Map<String, String>>();

    /**
     * 单例模式，获取DataUtil实例
     *
     * @return
     */
    public static DataUtil getInstance() {
        if (instance == null) {
            instance = new DataUtil();
        }
        return instance;
    }

    /**
     * 获取data文件夹绝对路径
     *
     * @return String
     */
    public String getPath() {
        String path = null;
        File directory = new File("data");// 设定为当前文件夹
        try {
            path = directory.getAbsolutePath();
        } catch (Exception e) {
            e.printStackTrace();
        }
        return path;
    }

    /**
     * 获取TestData文件路径
     *
     * @return String
     */
    public String getFilePath() {
        String filePath = this.getPath() + "\\TestData.xlsx";
        return filePath;
    }

    /**
     * 获取dataExcel工作文件
     *
     * @param filePath
     * @return HSSFWorkbook
     */
    public Workbook getHSSFWorkbook(String filePath) {
        Workbook wb = this.getWorkbook(filePath);
        return wb;
    }

    /**
     * 获取dataExcel工作文件
     *
     * @return HSSFWorkbook
     */
    public Workbook getHSSFWorkbook() {
        String filePath = this.getFilePath();
        Workbook wb = this.getWorkbook(filePath);
        return wb;
    }

    /**
     * 获取dataExcel工作文件
     *
     * @return HSSFWorkbook
     */
    public Workbook getWorkbook(String filePath) {
        FileInputStream fs = null;
        boolean isE2007 = false;
        Workbook wb = null;
        if (filePath.endsWith("xlsx")) {
            isE2007 = true;
        }
        InputStream input = null;  //建立输入流
        try {
            input = new FileInputStream(filePath);
            //根据文件格式(2003或者2007)来初始化
            if (isE2007) {
                wb = new XSSFWorkbook(input);
            } else {
                wb = new HSSFWorkbook(input);
            }
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException io) {
            io.printStackTrace();
        } finally {
            try {
                input.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
        return wb;
    }

    /**
     * 获取dataExcel表格
     *
     * @return String
     */
    public Sheet getSheet() {
        Workbook wb = this.getHSSFWorkbook();
        Sheet sheet = wb.getSheetAt(0);
        return sheet;
    }

    /**
     * 获取测试用例MAP  场景，MAP<用例名称,用例描述>
     *
     * @return
     */
    public HashMap<String, LinkedHashMap<String, String>> getTestCaseMap() {
        return testCaseMap;
    }

    /**
     * 设置测试用例MAP  场景，MAP<用例名称,用例描述>
     *
     * @param testCaseMap
     */
    public void setTestCaseMap(HashMap<String, LinkedHashMap<String, String>> testCaseMap) {
        this.testCaseMap = testCaseMap;
    }


    /**
     * 获取当前测试用例名称
     *
     * @return
     */
    public String getCurrentTestCase() {
        return currentTestCase;
    }

    /**
     * 设置当前测试用例为currentTestCase
     *
     * @param currentTestCase
     */
    public void setCurrentTestCase(String currentTestCase) {
        this.currentTestCase = currentTestCase;
    }

    /**
     * 获取场景名称
     *
     * @return
     */
    public String getSenario() {
        if (senario.equals(""))//默认取第一个
        {
            Map<String, String> stringStringMap = this.getSenarioList().get(0);
            senario = stringStringMap.keySet().iterator().next();
//            senario = GetSeniorData().get(0);
        }
        return senario;
    }

    /**
     * 设置场景名
     *
     * @param senario
     */
    public void setSenario(String senario) {
        this.senario = senario;
    }


    /**
     * 获取对应场景senario下，对应类名classname的用例名称
     *
     * @param senario
     * @param classname
     * @return
     */
    public String getTestCaseName(String senario, String classname) {
        String searched = "";
        HashMap<String, String> testCase = getTestCaseMap().get(senario);
        if (null != testCase) {
            Set<String> keys = testCase.keySet();
            for (String key : keys) {
                if (classname.equalsIgnoreCase(key)) {
                    searched = key;
                    break;
                }
            }
        }
        return searched;
    }

    /**
     * 获取对应场景senario下，对应用例名称casename的描述信息
     * 用例表中的第一列CaseId + ‘，’ + 第四列Description
     *
     * @param senario
     * @param caseName
     * @return
     */
    public String getTestCaseDesc(String senario, String caseName) {
        String desc = "";
        HashMap<String, String> testcase = getTestCaseMap().get(senario);
        if (null != testcase) {
            desc = testcase.get(caseName);
        }
        return desc;
    }

    String rowName;
    String columnName;
    String ddate;
    int i;
    int j;


    /**
     * 在TestData.xls文件中，根据columnName，获取所在的列号
     *
     * @param columnName
     * @return
     */
    public int getColumnCount(String columnName) {

        Sheet sheet = this.getSheet();
        Cell cell;
        for (int i = 0; i < 1; i++) {
            int j = 0;
            Row row = sheet.getRow(i);
            if (row != null) {
                sheet.getColumnBreaks();
                for (; j < row.getPhysicalNumberOfCells(); j++) {
                    cell = row.getCell(j);
                    if (cell != null) {
                        if (cell.toString().equals(columnName)) {
                            this.j = j;
                            break;
                        }
                    }
                    if (j == row.getPhysicalNumberOfCells() - 1) {
                        this.j = -1;
                    }
                }
            }
        }
        return this.j;
    }

    /**
     * 在TestData.xls文件中，根据rowname（场景名），获取所在的行号
     *
     * @param rowName 场景名
     * @return 行号
     */
    public int getRowCount(String rowName) {
        Sheet sheet = this.getSheet();
        Cell cell = null;
        for (int i = 1; i < sheet.getPhysicalNumberOfRows(); i++) {
            Row row = sheet.getRow(i); // 取得第 i Row
            if (row != null) {
                int j = 0;
                sheet.getColumnBreaks();
                for (; j < 1; j++) {
                    cell = row.getCell(j);
                    if (cell != null) {
                        if (cell.toString().equals(rowName)) {
                            this.i = i;
                            break;
                        }
                    }
                }
            }
        }
        return this.i;
    }


    /**
     * 在TestData.xls中，根据rowName（场景名称）和列名columName，获得对应单元格的值
     *
     * @param rowName
     * @param columnName
     * @return
     */
    public String GetData(String rowName, String columnName) {
        String data = null;
        String filePath;
        i = this.getRowCount(rowName);
        j = this.getColumnCount(columnName);
        if (j == -1) {
            return "";
        }
        Sheet sheet = this.getSheet();
        Cell cell;
        Row row = sheet.getRow(this.i); // 取得第 i Row
        if (row != null) {
            sheet.getColumnBreaks();
            cell = row.getCell(this.j);
            if (cell == null)
                data = "";
            else
                data = cell.toString();
        }
        return data;
    }

    /**
     * 在TestData.xls中，设置rowName（场景名称）和列名columnName对应的单元格的值
     *
     * @param rowName
     * @param columnName
     * @param ddate
     */
    public void SetData(String rowName, String columnName, String ddate) {
        i = this.getRowCount(rowName);
        j = this.getColumnCount(columnName);
        String filePath = this.getFilePath();
        Workbook wb = this.getHSSFWorkbook();
        try {
            Sheet sheet = wb.getSheetAt(0);
            Row row = sheet.getRow(this.i); // 取得第 i Row
            if (row != null) {
                // sheet.getColumnBreaks();
                Cell cell = row.getCell(this.j);
                if (cell == null) {
                    cell = row.createCell(this.j);
                }
                cell.setCellValue(ddate);
            }
            FileOutputStream fos = new FileOutputStream(filePath);
            fos.flush();
            wb.write(fos);
            fos.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    /**
     * 获取TestData.xls中第一个sheet的第一列，即所有场景名称，封装到 List<String>
     *
     * @return List<String>
     */
    public List<Map<String, String>> GetSeniorData() {
//        List<String> list = new ArrayList<String>();
        List<Map<String, String>> list = new ArrayList<Map<String, String>>();
        Sheet sheet = this.getSheet();
        try {
            Cell cell;
            for (int i = 1; i < sheet.getPhysicalNumberOfRows(); i++) {
                Row row = sheet.getRow(i); // 取得第 i Row
                if (row != null) {
                    sheet.getColumnBreaks();
                    if (row.getCell(0).toString() != "" && !"".equals(row.getCell(0).toString())) {
                        Map<String, String> seniorMap = new HashMap<String, String>();
                        String seniorName = row.getCell(0).toString();
                        String seniorProtocol = row.getCell(1).toString();
                        seniorMap.put(seniorName, seniorProtocol);
                        list.add(seniorMap);
//                        for (int j = 0; j < 1; j++) {
//                            cell = row.getCell(j);
//                            if ("".equals(cell.toString()))
//                                continue;
//                            list.add(cell.toString());
//                        }
                    }
                }
            }
        } catch (NullPointerException e) {
            new Throwable();
        }
        this.setSenarioList(list);
        return list;
    }

    /**
     * 获取baseDirName目录下，以targetFileName为名的文件，将文件的绝对路径add到filepath中
     *
     * @param baseDirName
     * @param targetFileName
     * @param filePath
     * @return
     */
    public String findFiles(String baseDirName, String targetFileName, List filePath) {

        File baseDir = new File(baseDirName); // 创建一个File对象
        if (!baseDir.exists() || !baseDir.isDirectory()) { // 判断目录是否存在
            System.out.println("文件查找失败：" + baseDirName + "不是一个目录！");
            logger.error("文件查找失败：" + baseDirName + "不是一个目录！");
        }
        String tempName = null; // 判断目录是否存在
        File tempFile;
        File[] files = baseDir.listFiles();
        for (int i = 0; i < files.length; i++) {
            tempFile = files[i];

            if (tempFile.isDirectory()) {
                findFiles(tempFile.getAbsolutePath(), targetFileName, filePath);
            } else if (tempFile.isFile()) {
                tempName = tempFile.getName();
                if (tempName.equalsIgnoreCase(targetFileName)) {
                    filePath.add(tempFile.getAbsolutePath());
                }
            }
        }
        return tempName;
    }

    /**
     * 根据filename，获取场景名对应的测试用例
     * 首先查找data目录下是否存在以filename为名的excel文件，如果不存在，则查找RunList.xls
     * 将excel中需要运行（RunFlag不为空且不为‘N’）的用例封装到LinkedHashMap<String, String>中，key为第二列TestCase，value为第一列+','+第四列描述
     *
     * @return LinkedHashMap<String, String>
     */
    public LinkedHashMap<String, String> GetTestCase(String filename) {

        LinkedHashMap<String, String> testCaseMap = new LinkedHashMap<String, String>();

        String caseFilePath;
        filename = filename + ".xlsx";
        List<String> filePath = new ArrayList<String>();
        findFiles(".\\data", filename, filePath);
        if (filePath.size() == 0) {
            findFiles(".\\data", "RunList.xlsx", filePath);
        }
        caseFilePath = filePath.get(0);  //获取第一个EXCEL,所以不能出现重名的
        System.out.println();
        try {
            Workbook wb = this.getHSSFWorkbook(caseFilePath);
            Sheet sheet = wb.getSheetAt(0);
            Cell cell;
            for (int i = 1; i < sheet.getPhysicalNumberOfRows(); i++) {
                Row row = sheet.getRow(i); // 取得第 i Row
                if (row != null) {
                    sheet.getColumnBreaks();
                    if (row.getCell(2) == null || "N".equals(row.getCell(2).toString()))
                        continue;
                    if (row.getCell(1) == null || "".equals(row.getCell(1).toString()))
                        break;
                    String testCase = row.getCell(1).toString().trim();
                    //刚刚修改 之前的value 是2
                    String testDesc = "";
                    try {
                        testDesc = row.getCell(0).toString().trim() + Constants.Comma + row.getCell(3).toString().trim();
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

    public List<Map<String, String>> getSenarioList() {
        return senarioList;
    }

    public void setSenarioList(List<Map<String, String>> senarioList) {
        this.senarioList = senarioList;
    }
}

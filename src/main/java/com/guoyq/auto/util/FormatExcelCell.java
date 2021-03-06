package com.guoyq.auto.util;

import org.apache.poi.ss.usermodel.Cell;

import java.text.DecimalFormat;

public class FormatExcelCell {
    //对excel表格中的double字段进行初始化
    public String formatExcelCelltoString(Cell cell){
        if (cell.getCellType()==cell.CELL_TYPE_NUMERIC){
            Double d = cell.getNumericCellValue();
            DecimalFormat df = new DecimalFormat("#.#########");
            return df.format(d);
        }else{
            return cell.toString();
        }
    }
}

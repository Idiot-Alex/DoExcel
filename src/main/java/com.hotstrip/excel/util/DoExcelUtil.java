package com.hotstrip.excel.util;

import com.hotstrip.enums.ExcelTypeEnums;
import com.hotstrip.exception.DoExcelException;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;
import java.io.InputStream;
import java.math.BigDecimal;

/**
 * poi 工作薄 相关工具类
 * 主要方法：
 * 创建一个工作薄
 * 创建一个 Sheet 表
 * 创建一个单元行
 * 创建一个单元格
 */
public class DoExcelUtil {

    // 默认 sheet 表名称
    private static String SHEET_TITLE = "Untitled";

    /**
     * 创建一个 工作薄
     * @param excelTypeEnums
     * @return
     * @throws IOException
     */
    public static Workbook createWorkBook(ExcelTypeEnums excelTypeEnums) {
        Workbook workbook;
        if(excelTypeEnums.equals(ExcelTypeEnums.XLSX))
            workbook = new XSSFWorkbook();
        else
            workbook = new HSSFWorkbook();
        return workbook;
    }

    /**
     * 创建 工作表
     * @param workbook
     * @param sheetTitle
     * @return
     */
    public static Sheet createSheet(Workbook workbook, String sheetTitle) {
        Sheet sheet = workbook.getSheet(sheetTitle != null ? sheetTitle : SHEET_TITLE);
        if (sheet == null) {
            sheet = workbook.createSheet(sheetTitle != null ? sheetTitle : SHEET_TITLE);
        }
        return sheet;
    }

    /**
     * 创建单元行
     * @param sheet
     * @param rowNum
     * @return
     */
    public static Row createRow(Sheet sheet, int rowNum) {
        return sheet.createRow(rowNum);
    }

    /**
     * 是否是数字
     * @param cellValue
     * @return
     */
    public static boolean isNum(Object cellValue) {
        return cellValue.toString().matches("^(-?\\d+)(\\.\\d+)?$");
    }

    /**
     * 是否是整数
     * @param cellValue
     * @return
     */
    public static boolean isIntger(Object cellValue) {
        return cellValue.toString().matches("^[-\\+]?[\\d]*$");
    }

    /**
     * 是否是百分数
     * @param cellValue
     * @return
     */
    public static boolean isPercent(Object cellValue) {
        return cellValue.toString().contains("%");
    }

    /**
     * 计算字符串长度
     * 1 个中文占 2 个字符位置
     * @param value
     * @return
     */
    public static int getValueLength(String value) {
        int valueLength = 0;
        for (int i = 0; i < value.length(); i++) {
            String temp = value.substring(i, i + 1);
            if (temp.matches("[\u4e00-\u9fa5]")) {
                valueLength += 2;
            } else {
                valueLength += 1;
            }
        }
        return valueLength;
    }

    /**
     * 根据字符长度计算单元格宽度
     * 这里实现不够精确
     * 具体参考下面链接 如何计算单元格宽度
     * https://blog.csdn.net/feg545/article/details/38661989
     */
    public static int calcCellWidth(int valueLength) {
        return (valueLength + 3) * Const.CELL_WIDTH;
    }
}

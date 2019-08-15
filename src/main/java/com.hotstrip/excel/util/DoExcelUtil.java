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
    public static Workbook createWorkBook(ExcelTypeEnums excelTypeEnums) throws IOException {
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
        return workbook.getSheet(sheetTitle != null ? sheetTitle : SHEET_TITLE);
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
}

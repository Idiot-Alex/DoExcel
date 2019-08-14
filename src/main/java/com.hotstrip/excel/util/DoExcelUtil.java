package com.hotstrip.excel.util;

import com.hotstrip.enums.ExcelTypeEnums;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;
import java.io.InputStream;

/**
 * poi 工作薄 相关工具类
 * 主要方法：
 * 创建一个工作薄
 * 创建一个 Sheet 表
 * 创建一个单元格
 */
public class DoExcelUtil {

    public static Workbook createWorkBook(InputStream inputStream, ExcelTypeEnums excelTypeEnums) throws IOException {
        Workbook workbook;
        // 检测输入流是否胃口
        if (inputStream == null) {

        }
        if(excelTypeEnums.equals(ExcelTypeEnums.XLSX))
            workbook = new XSSFWorkbook(inputStream);
        else
            workbook = new HSSFWorkbook(inputStream);
        return workbook;
    }
}

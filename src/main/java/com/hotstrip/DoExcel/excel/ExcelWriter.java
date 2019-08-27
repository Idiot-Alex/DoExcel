package com.hotstrip.DoExcel.excel;

import com.hotstrip.DoExcel.enums.ExcelTypeEnums;

import java.io.OutputStream;
import java.util.List;
import java.util.Locale;

/**
 * Created by Administrator on 2019/7/29.
 */
public class ExcelWriter {
    private ExcelContext excelContext;

    public ExcelWriter(OutputStream outputStream, ExcelTypeEnums excelTypeEnums) {
        excelContext = new ExcelContext(outputStream, excelTypeEnums);
    }

    /**
     * 写入数据
     * @param list 数据集合
     * @param clazz 实体类型
     * @return com.hotstrip.DoExcel.excel.ExcelWriter
     */
    public ExcelWriter write(List list, Class clazz) {
        excelContext.addContent(list, clazz);
        return this;
    }

    /**
     * 完成写入文件
     * @return com.hotstrip.DoExcel.excel.ExcelWriter
     */
    public ExcelWriter close() {
        excelContext.close();
        return this;
    }

    /**
     * 设置国际化资源
     * @param locale locale
     * @return com.hotstrip.DoExcel.excel.ExcelWriter
     */
    public ExcelWriter locale(Locale locale) {
        excelContext.setLocale(locale);
        return this;
    }

    /**
     * 单独写入行数据 支持国际化
     * @param list 数据集合
     * @return com.hotstrip.DoExcel.excel.ExcelWriter
     */
    public ExcelWriter writeRow(List<Object> list) {
        excelContext.addRow(list, true);
        return this;
    }
}

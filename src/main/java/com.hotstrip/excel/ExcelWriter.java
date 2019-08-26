package com.hotstrip.excel;

import com.hotstrip.enums.ExcelTypeEnums;

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
     * @param list
     * @param clazz
     * @return
     */
    public ExcelWriter write(List list, Class clazz) {
        excelContext.addContent(list, clazz);
        return this;
    }

    /**
     * 完成写入文件
     * @return
     */
    public ExcelWriter close() {
        excelContext.close();
        return this;
    }

    /**
     * 设置国际化资源
     * @param locale
     * @return
     */
    public ExcelWriter locale(Locale locale) {
        excelContext.setLocale(locale);
        return this;
    }

    /**
     * 单独写入行数据 支持国际化
     * @param list
     * @return
     */
    public ExcelWriter writeRow(List<Object> list) {
        excelContext.addRow(list, true);
        return this;
    }
}

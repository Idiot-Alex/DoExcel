package com.hotstrip.excel;

import com.hotstrip.enums.ExcelTypeEnums;

import java.io.OutputStream;
import java.util.List;

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
}

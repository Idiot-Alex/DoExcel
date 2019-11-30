package com.hotstrip.DoExcel.excel.eventmodel;

import java.io.FileInputStream;
import java.io.IOException;

/**
 * Created by Administrator on 2019/11/30.
 * 针对处理 Excel 读取操作做抽象
 * 至于区分 Excel 版本在实现类中
 */
public interface ExcelHandler {

    // 读取单元表
    void readSheet(FileInputStream fileInputStream) throws IOException;
}

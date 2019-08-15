package com.hotstrip.excel;

/**
 * Created by idiot on 2019/7/28.
 */
public class DoExcel {
    // excelConfig
    private ExcelContext excelConfig;

    public DoExcel() {
    }

    public ExcelContext getExcelConfig() {
        return excelConfig;
    }

    public void setExcelConfig(ExcelContext excelConfig) {
        this.excelConfig = excelConfig;
    }

    /**
     * ExcelWriter
     */
    public ExcelWriter getExcelWriter() {
        return ExcelWriter.builder()
                .title("test.xlsx")
                .build();
    }
}

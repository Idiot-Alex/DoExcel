package com.hotstrip.excel;

/**
 * Created by idiot on 2019/7/28.
 */
public class DoExcel {
    // excelConfig
    private ExcelConfig excelConfig;

    public DoExcel() {
    }

    public ExcelConfig getExcelConfig() {
        return excelConfig;
    }

    public void setExcelConfig(ExcelConfig excelConfig) {
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

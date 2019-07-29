package com.hotstrip.excel;

import org.apache.poi.ss.usermodel.Sheet;

import java.util.List;

/**
 * Created by Administrator on 2019/7/29.
 */
public class ExcelWriter {
    private String title;
    private List<Sheet> sheets;

    public String getTitle() {
        return title;
    }

    public void setTitle(String title) {
        this.title = title;
    }

    public List<Sheet> getSheets() {
        return sheets;
    }

    public void setSheets(List<Sheet> sheets) {
        this.sheets = sheets;
    }

    public static ExcelWriterBuilder builder() {
        return new ExcelWriterBuilder();
    }

    public ExcelWriter(String title, List<Sheet> sheets) {
        this.title = title;
        this.sheets = sheets;
    }

    public static class ExcelWriterBuilder {
        private String title;
        private List<Sheet> sheets;

        public ExcelWriterBuilder title(String title) {
            this.title = title;
            return this;
        }

        public ExcelWriterBuilder sheets(List<Sheet> sheets) {
            this.sheets = sheets;
            return this;
        }

        public ExcelWriter build() {
            return new ExcelWriter(title, sheets);
        }
    }
}

package com.hotstrip.excel;

import com.hotstrip.excel.util.DoExcelUtil;
import org.apache.poi.ss.usermodel.Sheet;

/**
 * Created by Administrator on 2019/7/29.
 */
public class ExcelWriter {
    private String title;
    private Sheet sheet;

    public String getTitle() {
        return title;
    }

    public void setTitle(String title) {
        this.title = title;
    }

    public Sheet getSheet() {
        return sheet;
    }

    public void setSheet(Sheet sheet) {
        this.sheet = sheet;
    }

    public static ExcelWriterBuilder builder() {
        return new ExcelWriterBuilder();
    }

    public ExcelWriter(String title, Sheet sheet) {
        this.title = title;
        this.sheet = sheet;
    }

    public static class ExcelWriterBuilder {
        private String title;
        private Sheet sheet;

        public ExcelWriterBuilder title(String title) {
            this.title = title;
            return this;
        }

        public ExcelWriterBuilder sheet(Sheet sheet) {
            this.sheet = sheet;
            return this;
        }

        public ExcelWriter build() {
            if (sheet == null) {
                sheet =
            }
            return new ExcelWriter(title, sheet);
        }
    }
}

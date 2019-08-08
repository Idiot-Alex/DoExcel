package com.hotstrip.excel;

import java.util.Locale;
import java.util.ResourceBundle;

/**
 * Excel config
 * such as Locale support
 */
public class ExcelConfig {
    // locale
    private Locale locale;
    // resourceBundle
    private ResourceBundle resourceBundle;

    public ExcelConfig() {
    }

    public Locale getLocale() {
        return locale;
    }

    public void setLocale(Locale locale) {
        this.locale = locale;
    }

    public ResourceBundle getResourceBundle() {
        return resourceBundle;
    }

    public void setResourceBundle(ResourceBundle resourceBundle) {
        this.resourceBundle = resourceBundle;
    }
}

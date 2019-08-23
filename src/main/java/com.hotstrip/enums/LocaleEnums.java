package com.hotstrip.enums;

import java.util.Locale;

/**
 * Created by idiot on 2017/5/27.
 */
public enum LocaleEnums {
    ZH_CN("简体中文", "zh_cn", Locale.SIMPLIFIED_CHINESE),
    EN_US("English US", "en_us", Locale.US),
    ZH_TW("繁體中文", "zh_tw", Locale.TRADITIONAL_CHINESE),
    ;

    private String name;
    private String value;
    private Locale locale;

    LocaleEnums(String name, String value, Locale locale) {
        this.name = name;
        this.value = value;
        this.locale = locale;
    }

    public String getName() {
        return name;
    }

    public String getValue() {
        return value;
    }

    public Locale getLocale() {
        return locale;
    }

    /******************************************************/

    public static Locale getLocaleByValue(String value){
        Locale locale = Locale.SIMPLIFIED_CHINESE;
        for (LocaleEnums code : LocaleEnums.values()){
            if(code.getValue().equals(value)){
                locale = code.getLocale();
                break;
            }
        }
        return locale;
    }
}

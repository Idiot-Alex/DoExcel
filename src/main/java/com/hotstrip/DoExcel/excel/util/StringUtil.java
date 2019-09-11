package com.hotstrip.DoExcel.excel.util;

/**
 * Created by Administrator on 2019/9/11.
 * 字符串操作类
 */
public class StringUtil {

    /**
     * 计算字符串长度
     * 1 个中文占 2 个字符位置
     * @param value String
     * @return int
     */
    public static int getValueLength(String value) {
        int valueLength = 0;
        for (int i = 0; i < value.length(); i++) {
            String temp = value.substring(i, i + 1);
            if (temp.matches("[\u4e00-\u9fa5]")) {
                valueLength += 2;
            } else {
                valueLength += 1;
            }
        }
        return valueLength;
    }

    /**
     * 去掉左边空格
     * @param s
     * @return
     */
    public static String ltrim(String s) {
        int i = 0;
        while (i < s.length() && Character.isWhitespace(s.charAt(i))) {
            i++;
        }
        return s.substring(i);
    }

    /**
     * 去掉右边空格
     * @param s
     * @return
     */
    public static String rtrim(String s) {
        int i = s.length()-1;
        while (i >= 0 && Character.isWhitespace(s.charAt(i))) {
            i--;
        }
        return s.substring(0, i+1);
    }
}

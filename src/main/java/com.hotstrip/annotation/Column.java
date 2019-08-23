package com.hotstrip.annotation;

import com.hotstrip.excel.util.Const;

import java.lang.annotation.*;

/**
 * 单个 sheet 表的元注解
 */
@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.FIELD)
@Documented
public @interface Column {
    // 字段名称
    String name();
    // 日期时间类型需要格式化
    String format() default Const.YMDHMS;
}

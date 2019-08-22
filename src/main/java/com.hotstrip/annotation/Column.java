package com.hotstrip.annotation;

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
}

package com.hotstrip.annotation;

import java.lang.annotation.*;

/**
 * 单个 sheet 表的元注解
 */
@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.FIELD)
@Documented
public @interface ColumnEnum {
    // 字段code
    String code();
    // 字段值
    String value();
}

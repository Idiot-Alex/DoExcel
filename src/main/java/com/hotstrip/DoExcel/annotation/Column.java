package com.hotstrip.DoExcel.annotation;

import java.lang.annotation.*;

/**
 * 单个字段的元注解
 */
@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.FIELD)
@Documented
public @interface Column {
    // 字段名称
    String name();

    /**
     * 重写字段数组注解
     * 比如某个字段 status 取值 {0: 无效, 1: 有效}
     * 最终导出数据用户希望看到的是字符串而非数字
     * 同时也支持国际化，只需要在实体类上加上 @DoSheet 注解，同时配置属性 localeResource
     * 然后资源文件对应 @ColumnEnum 的 value 属性即可
     * @return array
     */
    ColumnEnum[] columnEnums() default {};

    // 日期时间格式化
    String format() default "";
}

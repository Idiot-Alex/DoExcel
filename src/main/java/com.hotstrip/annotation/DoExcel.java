package com.hotstrip.annotation;

import java.lang.annotation.*;

/**
 * Created by idiot on 2019/6/3.
 */
@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.TYPE)
@Documented
public @interface DoExcel {
    String name() default "新建工作表.xlsx";
}

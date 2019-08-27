package com.hotstrip.annotation;

import java.lang.annotation.*;

/**
 * Created by idiot on 2019/6/3.
 */
@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.TYPE)
@Documented
public @interface DoSheet {
    // 标题
    String title() default "Untitled";
    // 国际化资源地址
    String localeResource() default "";
}

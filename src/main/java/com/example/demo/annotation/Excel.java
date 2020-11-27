package com.example.demo.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * @author liuhao
 */
@Retention(RetentionPolicy.RUNTIME)
@Target({ElementType.FIELD})
public @interface Excel {

    /**
     * 名称
     */
    String name() default "";

    /**
     * excel 列号
     */
    int column();

    /**
     * 是否必填
     */
    boolean isRequire() default false;

    /**
     * 匹配
     */
    String pattern() default RegExpStyle.NONE;

    /**
     * 解析错误原因
     */
    String cause() default "";
}

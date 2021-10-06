package com.example.comm;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * excel表数据类字段注解
 * 
 * @author Disen
 *
 */
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
public @interface Excel {
    /**
     * 列序号
     */
    int order() default -1;
    /**
     * 列名字
     */
    String[] name() default "";
    /**
     * 格式化
     */
    String fmt() default "";
}

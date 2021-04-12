package com.jz.zeus.excel.annotation;

import java.lang.annotation.*;

/**
 * @Author JZ
 * @Date 2021/4/8 17:31
 */
@Inherited
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
public @interface ValidationData {

    /**
     * 下拉框中的选项
     */
    String[] options();

    /**
     * 下拉框需要填充的行数
     */
    int rowNum() default 10000;

}

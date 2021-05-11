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
     * 将下拉框信息写入到单独sheet中
     */
    boolean asDicSheet() default false;

    /**
     * sheet 名称
     */
    String sheetName() default "";

    /**
     * 下拉框中的选项
     */
    String[] options();

    /**
     * 下拉框需要填充的行数
     */
    int rowNum() default 10000;

    /**
     * 是否校验数据是否为下拉框内容，若为false {@link #errorTitle} 和 {@link #errorMsg} 将无效
     */
    boolean checkDatavalidity() default true;

    /**
     * 错误信息box的标题
     */
    String errorTitle() default "";

    /**
     * 当输入的数据不是下拉框中数据时的提示信息
     */
    String errorMsg() default "";

}

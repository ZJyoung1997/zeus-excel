package com.jz.zeus.excel.annotation;

import java.awt.*;
import java.lang.annotation.*;

@Inherited
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
public @interface HeadColor {

    /**
     * 十六进制颜色码
     */
    String cellFillForegroundColor() default "";

    /**
     * 十六进制颜色码
     */
    String cellFillBackgroundColor() default "";

}

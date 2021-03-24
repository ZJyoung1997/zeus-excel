package com.jz.zeus.excel.validation;

import java.lang.annotation.*;

/**
 * @Author JZ
 * @Date 2021/3/24 19:08
 */
@Documented
@Constraint(
        validatedBy = {EmailValidator.class}
)
@Retention(RetentionPolicy.RUNTIME)
@Target({ElementType.METHOD, ElementType.FIELD, ElementType.ANNOTATION_TYPE, ElementType.CONSTRUCTOR, ElementType.PARAMETER})
public @interface Long {
}

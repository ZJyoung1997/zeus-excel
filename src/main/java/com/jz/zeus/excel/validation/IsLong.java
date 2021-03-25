package com.jz.zeus.excel.validation;

import jakarta.validation.Constraint;
import jakarta.validation.Payload;
import jakarta.validation.ReportAsSingleViolation;
import jakarta.validation.constraints.Pattern;

import java.lang.annotation.*;

/**
 * @Author JZ
 * @Date 2021/3/25 10:20
 */
@Documented
@Retention(RetentionPolicy.RUNTIME)
@ReportAsSingleViolation
@Constraint(validatedBy = {IsLongValidator.class})
@Target({ElementType.METHOD, ElementType.FIELD, ElementType.ANNOTATION_TYPE, ElementType.CONSTRUCTOR, ElementType.PARAMETER, ElementType.TYPE_USE})
public @interface IsLong {

    String message() default "value is not Long type";

    Class<?>[] groups() default {};

    Class<? extends Payload>[] payload() default {};

}

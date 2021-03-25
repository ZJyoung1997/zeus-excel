package com.jz.zeus.excel.validation;

import jakarta.validation.ConstraintValidator;
import jakarta.validation.ConstraintValidatorContext;

/**
 * @Author JZ
 * @Date 2021/3/25 13:48
 */
public class IsLongValidator implements ConstraintValidator<IsLong, Object> {

    private IsLong isLong;

    @Override
    public void initialize(IsLong constraintAnnotation) {
        this.isLong = constraintAnnotation;
    }

    @Override
    public boolean isValid(Object value, ConstraintValidatorContext constraintValidatorContext) {
        if (value == null) {
            return true;
        }
        if (value instanceof Long) {
            return true;
        }
        if (value instanceof String) {
            try {
                Long.valueOf((String) value);
                return true;
            } catch (NumberFormatException e) {}
        }
        return false;
    }

}

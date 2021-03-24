package com.jz.zeus.excel.util;

import jakarta.validation.ConstraintViolation;
import jakarta.validation.Validation;
import jakarta.validation.Validator;
import lombok.experimental.UtilityClass;
import org.hibernate.validator.HibernateValidator;

import java.util.Set;

/**
 * @author:JZ
 * @date:2021/3/24
 */
@UtilityClass
public class ValidatorUtils {

    private static final Validator validator = Validation.byProvider(HibernateValidator.class)
            .configure()
            .addProperty("hibernate.validator.fail_fast", "true")
            .buildValidatorFactory().getValidator();

    public <T> void validate(T bean) {
        if (bean == null) {
            return;
        }
        Set<ConstraintViolation<T>> constraintViolations = validator.validate(bean);
        constraintViolations.forEach(violation -> {
            violation.getMessage();
        });
    }

}

package com.jz.zeus.excel.util;

import jakarta.validation.ConstraintViolation;
import jakarta.validation.Validation;
import jakarta.validation.Validator;
import lombok.experimental.UtilityClass;
import org.hibernate.validator.HibernateValidator;

import java.util.*;

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

    /**
     * 校验bean中是否符合注解中的条件
     * @param bean 需校验对象
     * @param <T>
     * @return 错误信息，key 校验失败字段名、value 错误信息
     */
    public <T> Map<String, List<String>> validate(T bean) {
        if (bean == null) {
            return Collections.emptyMap();
        }
        Set<ConstraintViolation<T>> constraintViolations = validator.validate(bean);
        if (constraintViolations == null) {
            return Collections.emptyMap();
        }
        Map<String, List<String>> errorMessageMap = new HashMap<>();
        constraintViolations.forEach(violation -> {
            String fieldName = violation.getPropertyPath().toString();
            List<String> errorMessages = errorMessageMap.get(fieldName);
            if (errorMessages == null) {
                errorMessages = new ArrayList<>();
                errorMessageMap.put(fieldName, errorMessages);
            }
            errorMessages.add(violation.getMessage());
        });
        return errorMessageMap;
    }

}

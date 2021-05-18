package com.jz.zeus.excel.util;

import com.jz.zeus.excel.validator.VerifyResult;
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

    private static final Validator FAST_FAIL_VALIDATOR = Validation.byProvider(HibernateValidator.class)
            .configure()
            .failFast(true)
//            .addProperty("hibernate.validator.fail_fast", "true")
            .buildValidatorFactory().getValidator();

    private static final Validator VALIDATOR = Validation.byProvider(HibernateValidator.class)
            .configure()
            .buildValidatorFactory().getValidator();

    public <T> VerifyResult validate(T bean) {
        return validate(bean, false);
    }

    /**
     * 校验bean中是否符合注解中的条件
     * @param bean 需校验对象
     * @param <T>
     * @return 错误信息，key 校验失败字段名、value 错误信息
     */
    public <T> VerifyResult validate(T bean, boolean isFastFail) {
        if (bean == null) {
            return null;
        }
        Set<ConstraintViolation<T>> constraintViolations = (isFastFail ? FAST_FAIL_VALIDATOR : VALIDATOR).validate(bean);
        if (constraintViolations == null) {
            return null;
        }
        VerifyResult verifyResult = new VerifyResult();
        constraintViolations.forEach(violation -> {
            String fieldName = violation.getPropertyPath().toString();
            verifyResult.addErrorInfo(fieldName, violation.getMessage());
        });
        return verifyResult;
    }

}

package com.jz.zeus.excel.validator;

import cn.hutool.core.lang.Assert;
import cn.hutool.core.text.CharSequenceUtil;
import cn.hutool.core.util.StrUtil;
import com.jz.zeus.excel.util.ValidatorUtils;

import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.Objects;
import java.util.function.Predicate;

public class BeanValidator<T> {

    private final static String COLLECTION_FIELD_SUFFIX = "[%d].";

    private T bean;

    /**
     * 是否快速失败
     */
    private boolean enabledFastFail;

    private VerifyResult verifyResult;

    private List<VerifyInfo> verifyInfos = new ArrayList<>();

    private BeanValidator() {}

    private BeanValidator(T bean) {
        this.bean = bean;
    }

    public static <T> BeanValidator<T> build(T bean) {
        return new BeanValidator<>(bean);
    }

    public static <T> BeanValidator<T> build() {
        return new BeanValidator<>();
    }

    public BeanValidator<T> nonNull(Getter<T, Object> getter, String errorMsg) {
        return verify(getter, Objects::nonNull, errorMsg);
    }

    public BeanValidator<T> isNotBlank(Getter<T, String> getter, String errorMsg) {
        return verify(getter, CharSequenceUtil::isNotBlank, errorMsg);
    }

    public <R> BeanValidator<T> verify(Getter<T, R> getter, Predicate<R> condition, String errorMsg) {
        verifyInfos.add(new VerifyInfo(getter, condition, errorMsg));
        return this;
    }

    public <R> BeanValidator<T> verifyBean(Getter<T, R> getter, BeanValidator<R> beanValidator) {
        verifyInfos.add(new VerifyInfo(getter, beanValidator));
        return this;
    }

    public <R> BeanValidator<T> verifyByAnnotation(Getter<T, R> getter) {
        verifyInfos.add(new VerifyInfo(getter, true));
        return this;
    }

    public <E, R extends Iterable<E>> BeanValidator<T> verifyCollection(Getter<T, R> getter, BeanValidator<E> beanValidator) {
        verifyInfos.add(new VerifyInfo(getter, beanValidator));
        return this;
    }

    public VerifyResult doVerify(T bean) {
        this.bean = bean;
        return doVerify();
    }

    public VerifyResult doVerify() {
        Assert.notNull(this.bean, "Bean can not be null");
        verifyResult = new VerifyResult();
        for (VerifyInfo verifyInfo : verifyInfos) {
            String fieldName = verifyInfo.getter.getFieldName();
            Object value = verifyInfo.getter.apply(bean);
            if (verifyInfo.isAnnoationVerify) {
                verifyResult.addVerifyResult(ValidatorUtils.validate(value, enabledFastFail), fieldName, StrUtil.DOT);
            }
            if (Objects.isNull(verifyInfo.childValidator)) {
                if (Objects.nonNull(verifyInfo.condition) && !verifyInfo.condition.test(value)) {
                    verifyResult.addErrorInfo(fieldName, verifyInfo.errorMsg);
                }
                continue;
            }
            BeanValidator childValidator = verifyInfo.childValidator;
            if (value instanceof Iterable) {
                int index = 0;
                Iterator<?> iterator = ((Iterable<?>) value).iterator();
                while (iterator.hasNext()) {
                    Object element = iterator.next();
                    verifyResult.addVerifyResult(childValidator.doVerify(element), fieldName,
                            String.format(COLLECTION_FIELD_SUFFIX, index++));
                }
            } else {
                verifyResult.addVerifyResult(childValidator.doVerify(value), fieldName, StrUtil.DOT);
            }
        }
        return verifyResult;
    }

    private class VerifyInfo {

        private boolean isAnnoationVerify;

        private Getter getter;

        private Predicate condition;

        private BeanValidator childValidator;

        private String errorMsg;

        private VerifyInfo(Getter getter, boolean isAnnoationVerify) {
            this.getter = getter;
            this.isAnnoationVerify = isAnnoationVerify;
        }

        private VerifyInfo(Getter getter, BeanValidator childValidator) {
            Assert.notNull(getter, "Getter can not be null");
            Assert.notNull(childValidator, "BeanValidator can not be null");
            this.getter = getter;
            this.childValidator = childValidator;
        }

        private VerifyInfo(Getter getter, Predicate condition, String errorMsg) {
            Assert.notNull(getter, "Getter can not be null");
            Assert.notNull(condition, "Condition can not be null");
            this.getter = getter;
            this.condition = condition;
            this.errorMsg = errorMsg;
        }

    }

}

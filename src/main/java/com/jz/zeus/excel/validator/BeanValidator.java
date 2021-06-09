package com.jz.zeus.excel.validator;

import cn.hutool.core.lang.Assert;
import cn.hutool.core.text.CharSequenceUtil;
import cn.hutool.core.util.StrUtil;
import com.jz.zeus.excel.interfaces.Getter;
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

    public BeanValidator<T> nonNull(boolean precondition, Getter<T, Object> getter, String errorMsg) {
        return verify(precondition, getter, Objects::nonNull, errorMsg);
    }

    public BeanValidator<T> nonNull(Predicate<T> precondition, Getter<T, Object> getter, String errorMsg) {
        return verify(precondition, getter, Objects::nonNull, errorMsg);
    }

    public BeanValidator<T> isNotBlank(Getter<T, String> getter, String errorMsg) {
        return verify(getter, CharSequenceUtil::isNotBlank, errorMsg);
    }

    public BeanValidator<T> isNotBlank(boolean precondition, Getter<T, String> getter, String errorMsg) {
        return verify(precondition, getter, CharSequenceUtil::isNotBlank, errorMsg);
    }

    public BeanValidator<T> isNotBlank(Predicate<T> precondition, Getter<T, String> getter, String errorMsg) {
        return verify(precondition, getter, CharSequenceUtil::isNotBlank, errorMsg);
    }

    public <R> BeanValidator<T> verify(Getter<T, R> getter, Predicate<R> condition, String errorMsg) {
        verifyInfos.add(VerifyInfo.build(getter, condition, errorMsg));
        return this;
    }

    public <R> BeanValidator<T> verify(boolean precondition, Getter<T, R> getter, Predicate<R> condition, String errorMsg) {
        verifyInfos.add(VerifyInfo.build(precondition, getter, condition, errorMsg));
        return this;
    }

    public <R> BeanValidator<T> verify(Predicate<T> precondition, Getter<T, R> getter, Predicate<R> condition, String errorMsg) {
        verifyInfos.add(VerifyInfo.build(precondition, getter, condition, errorMsg));
        return this;
    }

    public <R> BeanValidator<T> verifyBean(Getter<T, R> getter, BeanValidator<R> beanValidator) {
        verifyInfos.add(VerifyInfo.build(getter, beanValidator));
        return this;
    }

    public <R> BeanValidator<T> verifyBean(boolean precondition, Getter<T, R> getter, BeanValidator<R> beanValidator) {
        verifyInfos.add(VerifyInfo.build(precondition, getter, beanValidator));
        return this;
    }

    public <R> BeanValidator<T> verifyBean(Predicate<T> precondition, Getter<T, R> getter, BeanValidator<R> beanValidator) {
        verifyInfos.add(VerifyInfo.build(precondition, getter, beanValidator));
        return this;
    }

    public <R> BeanValidator<T> verifyByAnnotation(Getter<T, R> getter) {
        verifyInfos.add(VerifyInfo.build(getter, true));
        return this;
    }

    public <R> BeanValidator<T> verifyByAnnotation(boolean precondition, Getter<T, R> getter) {
        verifyInfos.add(VerifyInfo.build(precondition, getter, true));
        return this;
    }

    public <R> BeanValidator<T> verifyByAnnotation(Predicate<T> precondition, Getter<T, R> getter) {
        verifyInfos.add(VerifyInfo.build(precondition, getter, true));
        return this;
    }

    public <E, R extends Iterable<E>> BeanValidator<T> verifyCollection(Getter<T, R> getter, BeanValidator<E> beanValidator) {
        verifyInfos.add(VerifyInfo.build(getter, beanValidator));
        return this;
    }

    public <E, R extends Iterable<E>> BeanValidator<T> verifyCollection(Predicate<T> precondition, Getter<T, R> getter, BeanValidator<E> beanValidator) {
        verifyInfos.add(VerifyInfo.build(precondition, getter, beanValidator));
        return this;
    }

    public <E, R extends Iterable<E>> BeanValidator<T> verifyCollection(boolean precondition, Getter<T, R> getter, BeanValidator<E> beanValidator) {
        verifyInfos.add(VerifyInfo.build(precondition, getter, beanValidator));
        return this;
    }

    public VerifyResult doVerify(T bean, boolean enabledFastFail) {
        this.bean = bean;
        return doVerify(enabledFastFail);
    }

    public VerifyResult doVerify() {
        return doVerify(false);
    }

    public VerifyResult doVerify(boolean enabledFastFail) {
        Assert.notNull(this.bean, "Bean can not be null");
        this.enabledFastFail = enabledFastFail;
        verifyResult = new VerifyResult();
        for (VerifyInfo verifyInfo : verifyInfos) {
            if (!verifyInfo.verifyPrecondition(bean)) {
                continue;
            }
            Getter getter = verifyInfo.getGetter();
            String fieldName = getter.getFieldName();
            Object value = getter.apply(bean);
            if (verifyInfo.isAnnoationVerify()) {
                verifyResult.addVerifyResult(ValidatorUtils.validate(value, enabledFastFail), fieldName, StrUtil.DOT);
                if (stopVerify()) {
                    return verifyResult;
                }
            }
            if (Objects.isNull(verifyInfo.getChildValidator())) {
                if (Objects.nonNull(verifyInfo.getCondition()) && !verifyInfo.getCondition().test(value)) {
                    verifyResult.addErrorInfo(fieldName, verifyInfo.getErrorMsg());
                }
                if (stopVerify()) {
                    return verifyResult;
                }
                continue;
            }
            BeanValidator childValidator = verifyInfo.getChildValidator();
            if (value instanceof Iterable) {
                int index = 0;
                Iterator<?> iterator = ((Iterable<?>) value).iterator();
                while (iterator.hasNext()) {
                    Object element = iterator.next();
                    verifyResult.addVerifyResult(childValidator.doVerify(element, enabledFastFail), fieldName,
                            String.format(COLLECTION_FIELD_SUFFIX, index++));
                }
            } else {
                verifyResult.addVerifyResult(childValidator.doVerify(value, enabledFastFail), fieldName, StrUtil.DOT);
            }
            if (stopVerify()) {
                return verifyResult;
            }
        }
        return verifyResult;
    }

    private boolean stopVerify() {
        return enabledFastFail && verifyResult.hasError();
    }

}

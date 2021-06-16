package com.jz.zeus.excel.validator;

import cn.hutool.core.lang.Assert;
import cn.hutool.core.text.CharSequenceUtil;
import cn.hutool.core.util.StrUtil;
import com.jz.zeus.excel.interfaces.FieldGetter;
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

    public BeanValidator<T> nonNull(FieldGetter<T, Object> fieldGetter, String errorMsg) {
        return verify(fieldGetter, Objects::nonNull, errorMsg);
    }

    public BeanValidator<T> nonNull(boolean precondition, FieldGetter<T, Object> fieldGetter, String errorMsg) {
        return verify(precondition, fieldGetter, Objects::nonNull, errorMsg);
    }

    public BeanValidator<T> nonNull(Predicate<T> precondition, FieldGetter<T, Object> fieldGetter, String errorMsg) {
        return verify(precondition, fieldGetter, Objects::nonNull, errorMsg);
    }

    public BeanValidator<T> isNotBlank(FieldGetter<T, String> fieldGetter, String errorMsg) {
        return verify(fieldGetter, CharSequenceUtil::isNotBlank, errorMsg);
    }

    public BeanValidator<T> isNotBlank(boolean precondition, FieldGetter<T, String> fieldGetter, String errorMsg) {
        return verify(precondition, fieldGetter, CharSequenceUtil::isNotBlank, errorMsg);
    }

    public BeanValidator<T> isNotBlank(Predicate<T> precondition, FieldGetter<T, String> fieldGetter, String errorMsg) {
        return verify(precondition, fieldGetter, CharSequenceUtil::isNotBlank, errorMsg);
    }

    public <R> BeanValidator<T> verify(FieldGetter<T, R> fieldGetter, Predicate<R> condition, String errorMsg) {
        verifyInfos.add(VerifyInfo.build(fieldGetter, condition, errorMsg));
        return this;
    }

    public <R> BeanValidator<T> verify(boolean precondition, FieldGetter<T, R> fieldGetter, Predicate<R> condition, String errorMsg) {
        verifyInfos.add(VerifyInfo.build(precondition, fieldGetter, condition, errorMsg));
        return this;
    }

    public <R> BeanValidator<T> verify(Predicate<T> precondition, FieldGetter<T, R> fieldGetter, Predicate<R> condition, String errorMsg) {
        verifyInfos.add(VerifyInfo.build(precondition, fieldGetter, condition, errorMsg));
        return this;
    }

    public <R> BeanValidator<T> verifyBean(FieldGetter<T, R> fieldGetter, BeanValidator<R> beanValidator) {
        verifyInfos.add(VerifyInfo.build(fieldGetter, beanValidator));
        return this;
    }

    public <R> BeanValidator<T> verifyBean(boolean precondition, FieldGetter<T, R> fieldGetter, BeanValidator<R> beanValidator) {
        verifyInfos.add(VerifyInfo.build(precondition, fieldGetter, beanValidator));
        return this;
    }

    public <R> BeanValidator<T> verifyBean(Predicate<T> precondition, FieldGetter<T, R> fieldGetter, BeanValidator<R> beanValidator) {
        verifyInfos.add(VerifyInfo.build(precondition, fieldGetter, beanValidator));
        return this;
    }

    public <R> BeanValidator<T> verifyByAnnotation(FieldGetter<T, R> fieldGetter) {
        verifyInfos.add(VerifyInfo.build(fieldGetter, true));
        return this;
    }

    public <R> BeanValidator<T> verifyByAnnotation(boolean precondition, FieldGetter<T, R> fieldGetter) {
        verifyInfos.add(VerifyInfo.build(precondition, fieldGetter, true));
        return this;
    }

    public <R> BeanValidator<T> verifyByAnnotation(Predicate<T> precondition, FieldGetter<T, R> fieldGetter) {
        verifyInfos.add(VerifyInfo.build(precondition, fieldGetter, true));
        return this;
    }

    public <E, R extends Iterable<E>> BeanValidator<T> verifyCollection(FieldGetter<T, R> fieldGetter, BeanValidator<E> beanValidator) {
        verifyInfos.add(VerifyInfo.build(fieldGetter, beanValidator));
        return this;
    }

    public <E, R extends Iterable<E>> BeanValidator<T> verifyCollection(Predicate<T> precondition, FieldGetter<T, R> fieldGetter, BeanValidator<E> beanValidator) {
        verifyInfos.add(VerifyInfo.build(precondition, fieldGetter, beanValidator));
        return this;
    }

    public <E, R extends Iterable<E>> BeanValidator<T> verifyCollection(boolean precondition, FieldGetter<T, R> fieldGetter, BeanValidator<E> beanValidator) {
        verifyInfos.add(VerifyInfo.build(precondition, fieldGetter, beanValidator));
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
            FieldGetter fieldGetter = verifyInfo.getFieldGetter();
            String fieldName = fieldGetter.getFieldName();
            Object value = fieldGetter.apply(bean);
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

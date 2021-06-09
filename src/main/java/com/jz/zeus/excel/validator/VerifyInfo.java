package com.jz.zeus.excel.validator;

import cn.hutool.core.lang.Assert;
import com.jz.zeus.excel.interfaces.Getter;
import lombok.AccessLevel;

import java.util.function.Predicate;

/**
 * @Author JZ
 * @Date 2021/6/9 10:48
 */
@lombok.Getter(AccessLevel.PACKAGE)
class VerifyInfo {

    private boolean isAnnoationVerify;

    private Getter getter;

    private Predicate condition;

    /**
     * 前置条件
     */
    private Boolean precondition;

    /**
     * 前置条件，与 {@link #precondition}同时只能存在一个
     */
    private Predicate preconditionPredicate;

    private BeanValidator childValidator;

    private String errorMsg;

    private VerifyInfo() {}

    public static VerifyInfo build(Getter getter, boolean isAnnoationVerify) {
        VerifyInfo verifyInfo = new VerifyInfo();
        verifyInfo.getter = getter;
        verifyInfo.isAnnoationVerify = isAnnoationVerify;
        return verifyInfo;
    }

    public static VerifyInfo build(boolean precondition, Getter getter, boolean isAnnoationVerify) {
        Assert.notNull(getter, "Getter can not be null");
        VerifyInfo verifyInfo = new VerifyInfo();
        verifyInfo.precondition = precondition;
        verifyInfo.getter = getter;
        verifyInfo.isAnnoationVerify = isAnnoationVerify;
        return verifyInfo;
    }

    public static VerifyInfo build(Predicate precondition, Getter getter, boolean isAnnoationVerify) {
        Assert.notNull(getter, "Getter can not be null");
        Assert.notNull(getter, "Precondition can not be null");
        VerifyInfo verifyInfo = new VerifyInfo();
        verifyInfo.preconditionPredicate = precondition;
        verifyInfo.getter = getter;
        verifyInfo.isAnnoationVerify = isAnnoationVerify;
        return verifyInfo;
    }

    public static VerifyInfo build(Getter getter, BeanValidator childValidator) {
        Assert.notNull(getter, "Getter can not be null");
        Assert.notNull(childValidator, "BeanValidator can not be null");
        VerifyInfo verifyInfo = new VerifyInfo();
        verifyInfo.getter = getter;
        verifyInfo.childValidator = childValidator;
        return verifyInfo;
    }

    public static VerifyInfo build(boolean precondition, Getter getter, BeanValidator childValidator) {
        Assert.notNull(getter, "Getter can not be null");
        Assert.notNull(childValidator, "BeanValidator can not be null");
        VerifyInfo verifyInfo = new VerifyInfo();
        verifyInfo.precondition = precondition;
        verifyInfo.getter = getter;
        verifyInfo.childValidator = childValidator;
        return verifyInfo;
    }

    public static VerifyInfo build(Predicate precondition, Getter getter, BeanValidator childValidator) {
        Assert.notNull(getter, "Getter can not be null");
        Assert.notNull(childValidator, "BeanValidator can not be null");
        Assert.notNull(getter, "Precondition can not be null");
        VerifyInfo verifyInfo = new VerifyInfo();
        verifyInfo.preconditionPredicate = precondition;
        verifyInfo.getter = getter;
        verifyInfo.childValidator = childValidator;
        return verifyInfo;
    }

    public static VerifyInfo build(Getter getter, Predicate condition, String errorMsg) {
        Assert.notNull(getter, "Getter can not be null");
        Assert.notNull(condition, "Condition can not be null");
        VerifyInfo verifyInfo = new VerifyInfo();
        verifyInfo.getter = getter;
        verifyInfo.condition = condition;
        verifyInfo.errorMsg = errorMsg;
        return verifyInfo;
    }

    public static VerifyInfo build(boolean precondition, Getter getter, Predicate condition, String errorMsg) {
        Assert.notNull(getter, "Getter can not be null");
        Assert.notNull(condition, "Condition can not be null");
        VerifyInfo verifyInfo = new VerifyInfo();
        verifyInfo.precondition = precondition;
        verifyInfo.getter = getter;
        verifyInfo.condition = condition;
        verifyInfo.errorMsg = errorMsg;
        return verifyInfo;
    }

    public static VerifyInfo build(Predicate precondition, Getter getter, Predicate condition, String errorMsg) {
        Assert.notNull(getter, "Getter can not be null");
        Assert.notNull(condition, "Condition can not be null");
        Assert.notNull(getter, "Precondition can not be null");
        VerifyInfo verifyInfo = new VerifyInfo();
        verifyInfo.preconditionPredicate = precondition;
        verifyInfo.getter = getter;
        verifyInfo.condition = condition;
        verifyInfo.errorMsg = errorMsg;
        return verifyInfo;
    }

    public boolean verifyPrecondition(Object bean) {
        if (precondition == null && preconditionPredicate == null) {
            return true;
        }
        if (precondition != null) {
            return precondition;
        } else if (preconditionPredicate != null) {
            return preconditionPredicate.test(bean);
        }
        return false;
    }

}

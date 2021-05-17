package com.jz.zeus.excel.validator;

import java.util.Collection;
import java.util.Map;

public class BeanValidator<T> {

    private Map<String, String> errorInfoMap;

    public <R> BeanValidator<T> verifyCollection(Getter<T, Collection<R>> function, BeanValidator<R> validator) {
        return this;
    }

    public BeanValidator<T> nonNull(Getter<T, String> getter) {
        getter.getFieldName();
        return this;
    }

    public static <T> BeanValidator<T> build(T data) {
        return new BeanValidator<>();
    }

}

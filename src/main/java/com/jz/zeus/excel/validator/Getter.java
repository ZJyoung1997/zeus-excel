package com.jz.zeus.excel.validator;

import lombok.SneakyThrows;

import java.lang.invoke.SerializedLambda;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.util.function.Function;

public interface Getter<T, R> extends Function<T, R> {

    @SneakyThrows
    default String getFieldName() {
        Method method = this.getClass().getDeclaredMethod("writeReplace");
        method.setAccessible(true);
        SerializedLambda serializedLambda = (SerializedLambda) method.invoke(this);
        String fieldName = serializedLambda.getImplMethodName().substring("get".length());
        fieldName = fieldName.replaceFirst(fieldName.charAt(0) + "", (fieldName.charAt(0) + "").toLowerCase());
        Field field = Class.forName(serializedLambda.getImplClass().replace("/", "."))
                .getDeclaredField(fieldName);
        return fieldName;
    }

}

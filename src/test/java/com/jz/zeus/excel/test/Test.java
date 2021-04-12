package com.jz.zeus.excel.test;

import com.jz.zeus.excel.util.UnsafeFieldAccessor;
import lombok.Data;
import lombok.SneakyThrows;

import java.lang.reflect.Field;
import java.lang.reflect.Method;

/**
 * @Author JZ
 * @Date 2021/4/12 16:31
 */
public class Test {

    @Data
    static class T {
        private Object obj;
    }

    @SneakyThrows
    public static void main(String[] args) {
        Method getMethod = T.class.getMethod("getObj");
        Method setMethod = T.class.getMethod("setObj", Object.class);
        Field field = T.class.getDeclaredField("obj");
        field.setAccessible(true);
        UnsafeFieldAccessor accessor = new UnsafeFieldAccessor(field);
        T t = new T();
        Object v = new Object();
        int n = 10000000;
        long start;


        start = System.currentTimeMillis();
        for (int i = 0; i < n; i++) {
            t = new T();
            t.setObj(new Object());
            accessor.getObject(t);
        }
        System.out.printf("unsafe get: %d\n", System.currentTimeMillis() - start);

        start = System.currentTimeMillis();
        for (int i = 0; i < n; i++) {
            field.get(t);
        }
        System.out.printf("reflect get: %d\n", System.currentTimeMillis() - start);

        start = System.currentTimeMillis();
        for (int i = 0; i < n; i++) {
            getMethod.invoke(t);
        }
        System.out.printf("cache method get: %d\n", System.currentTimeMillis() - start);

        // -----------------------------------------------------

        start = System.currentTimeMillis();
        for (int i = 0; i < n; i++) {
            accessor.setObject(t, v);
        }
        System.out.printf("unsafe set: %d\n", System.currentTimeMillis() - start);

        start = System.currentTimeMillis();
        for (int i = 0; i < n; i++) {
            field.set(t, v);
        }
        System.out.printf("reflect set: %d\n", System.currentTimeMillis() - start);

        start = System.currentTimeMillis();
        for (int i = 0; i < n; i++) {
            setMethod.invoke(t, v);
        }
        System.out.printf("cache method set: %d\n", System.currentTimeMillis() - start);
    }

}

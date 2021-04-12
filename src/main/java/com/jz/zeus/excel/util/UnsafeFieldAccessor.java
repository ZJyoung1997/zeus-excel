package com.jz.zeus.excel.util;

import sun.misc.Unsafe;

import java.lang.reflect.Field;

/**
 * @Author JZ
 * @Date 2021/4/12 16:30
 */
public class UnsafeFieldAccessor {

    private static final Unsafe unsafe ;

    static {
        try {
            Field field = Unsafe.class.getDeclaredField("theUnsafe");
            field.setAccessible(true);
            unsafe = (Unsafe) field.get(null);
        } catch (Exception e) {
            throw new IllegalStateException(e);
        }
    }

    private final long fieldOffset;

    public UnsafeFieldAccessor(Field field) {
        fieldOffset = unsafe.objectFieldOffset(field);
    }

    public Object getObject(Object object) {
        return unsafe.getObject(object, fieldOffset);
    }

    public void setObject(Object object, Object value) {
        unsafe.putObject(object, fieldOffset, value);
    }

}

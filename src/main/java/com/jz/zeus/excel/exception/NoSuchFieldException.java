package com.jz.zeus.excel.exception;

/**
 * @Author JZ
 * @Date 2021/6/3 10:36
 */
public class NoSuchFieldException extends RuntimeException {

    public NoSuchFieldException() {}

    public NoSuchFieldException(String message) {
        super(message);
    }
}

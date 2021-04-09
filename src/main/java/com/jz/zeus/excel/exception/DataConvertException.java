package com.jz.zeus.excel.exception;

import lombok.Getter;

/**
 * 数据转换异常，用于在 {@link com.alibaba.excel.converters.Converter} 中抛出信息
 * @Author JZ
 * @Date 2021/4/9 11:41
 */
public class DataConvertException extends RuntimeException {

    @Getter
    private String errorMsg;

    public DataConvertException(String errorMsg) {
        this.errorMsg = errorMsg;
    }

    public DataConvertException(String errorMsg, Throwable cause) {
        super(cause);
        this.errorMsg = errorMsg;
    }

}

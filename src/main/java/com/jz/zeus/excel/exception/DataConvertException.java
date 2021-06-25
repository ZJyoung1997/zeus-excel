package com.jz.zeus.excel.exception;

import com.jz.zeus.excel.CellErrorInfo;
import lombok.Getter;

import java.util.List;

/**
 * 数据转换异常，用于在 {@link com.alibaba.excel.converters.Converter} 中抛出信息
 * @Author JZ
 * @Date 2021/4/9 11:41
 */
public class DataConvertException extends RuntimeException {

    @Getter
    private List<CellErrorInfo> cellErrorInfos;

    public DataConvertException(List<CellErrorInfo> cellErrorInfos) {
        this.cellErrorInfos = cellErrorInfos;
    }

    public DataConvertException(String errorMsg) {
        super(errorMsg);
    }

    public DataConvertException(String errorMsg, Throwable cause) {
        super(errorMsg, cause);
    }

}

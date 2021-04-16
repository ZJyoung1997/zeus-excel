package com.jz.zeus.excel;

import cn.hutool.core.collection.CollUtil;
import cn.hutool.core.lang.Assert;
import cn.hutool.core.util.ArrayUtil;
import cn.hutool.core.util.StrUtil;
import lombok.Data;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collection;

/**
 * @Author JZ
 * @Date 2021/3/22 17:56
 */
@Data
public class CellErrorInfo {

    private Integer rowIndex;

    private Integer columnIndex;

    private String fieldName;

    private String headName;

    private Collection<String> errorMsgs;

    private CellErrorInfo() {}

    public static CellErrorInfo buildByHead(int rowIndex, String headName, String... errorMsgs) {
        return build(rowIndex, null, headName, null, errorMsgs);
    }

    public static CellErrorInfo buildByField(int rowIndex, String fieldName, String... errorMsgs) {
        return build(rowIndex, fieldName, null, null, errorMsgs);
    }

    public static CellErrorInfo buildByColumnIndex(int rowIndex, int columnIndex, String... errorMsgs) {
        return build(rowIndex, null, null, columnIndex, errorMsgs);
    }

    public static CellErrorInfo build(int rowIndex, String fieldName, String headName, Integer columnIndex, String... errorMsgs) {
        Assert.isTrue(rowIndex >= 0, "RowIndex has to be greater than or equal to 0");
        Assert.isFalse(columnIndex == null && StrUtil.isBlank(fieldName) && StrUtil.isBlank(headName),
                "ColumnIndex and FieldName and HeadName not all empty");
        CellErrorInfo errorInfo = new CellErrorInfo();
        errorInfo.setRowIndex(rowIndex);
        if (columnIndex != null) {
            Assert.isTrue(columnIndex >= 0, "ColumnIndex has to be greater than or equal to 0");
            errorInfo.setColumnIndex(columnIndex);
        }
        errorInfo.setFieldName(fieldName);
        errorInfo.setHeadName(headName);
        if (ArrayUtil.isNotEmpty(errorMsgs)) {
            errorInfo.setErrorMsgs(new ArrayList<>(Arrays.asList(errorMsgs)));
        } else {
            errorInfo.setErrorMsgs(new ArrayList<>());
        }
        return errorInfo;
    }

    public CellErrorInfo addErrorMsg(String msg) {
        if (StrUtil.isNotBlank(msg)) {
            errorMsgs.add(msg);
        }
        return this;
    }

    public CellErrorInfo addErrorMsg(Collection<String> msgs) {
        if (CollUtil.isNotEmpty(msgs)) {
            errorMsgs.addAll(msgs);
        }
        return this;
    }

    public boolean hasError() {
        return CollUtil.isNotEmpty(errorMsgs);
    }

}

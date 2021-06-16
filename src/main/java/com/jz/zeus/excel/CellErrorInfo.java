package com.jz.zeus.excel;

import cn.hutool.core.collection.CollUtil;
import cn.hutool.core.lang.Assert;
import cn.hutool.core.util.ArrayUtil;
import cn.hutool.core.util.StrUtil;
import com.jz.zeus.excel.interfaces.FieldGetter;
import lombok.Data;
import lombok.SneakyThrows;
import lombok.experimental.Accessors;

import java.util.ArrayList;
import java.util.Collection;

/**
 * @Author JZ
 * @Date 2021/3/22 17:56
 */
@Data
@Accessors(chain = true)
public class CellErrorInfo implements Cloneable {

    private Integer rowIndex;

    private Integer columnIndex;

    private String fieldName;

    private String headName;

    private Collection<String> errorMsgs;

    private CellErrorInfo() {}

    public static CellErrorInfo buildByHead(int rowIndex, String headName, String... errorMsgs) {
        return build(rowIndex, null, headName, null, errorMsgs);
    }

    public static CellErrorInfo buildByHead(int rowIndex, String headName,  Collection<String> errorMsgs) {
        return build(rowIndex, null, headName, null, errorMsgs);
    }

    public static CellErrorInfo buildByField(int rowIndex, String fieldName, String... errorMsgs) {
        return build(rowIndex, fieldName, null, null, errorMsgs);
    }

    public static <T, R> CellErrorInfo buildByField(int rowIndex, FieldGetter<T, R> fieldGetter, String... errorMsgs) {
        return build(rowIndex, fieldGetter.getFieldName(), null, null, errorMsgs);
    }

    public static CellErrorInfo buildByField(int rowIndex, String fieldName,  Collection<String> errorMsgs) {
        return build(rowIndex, fieldName, null, null, errorMsgs);
    }

    public static <T, R> CellErrorInfo buildByField(int rowIndex, FieldGetter<T, R> fieldGetter,  Collection<String> errorMsgs) {
        return build(rowIndex, fieldGetter.getFieldName(), null, null, errorMsgs);
    }

    public static CellErrorInfo buildByColumnIndex(int rowIndex, int columnIndex, String... errorMsgs) {
        return build(rowIndex, null, null, columnIndex, errorMsgs);
    }

    public static CellErrorInfo buildByColumnIndex(int rowIndex, int columnIndex,  Collection<String> errorMsgs) {
        return build(rowIndex, null, null, columnIndex, errorMsgs);
    }

    private static CellErrorInfo build(int rowIndex, String fieldName, String headName, Integer columnIndex, String... errorMsgs) {
        if (ArrayUtil.isNotEmpty(errorMsgs)) {
            return build(rowIndex, fieldName, headName, columnIndex, CollUtil.toList(errorMsgs));
        }
        return build(rowIndex, fieldName, headName, columnIndex, (Collection<String>) null);
    }

    private static CellErrorInfo build(int rowIndex, String fieldName, String headName, Integer columnIndex, Collection<String> errorMsgs) {
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
        if (errorMsgs != null) {
            errorInfo.setErrorMsgs(errorMsgs);
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

    @Override
    @SneakyThrows
    public CellErrorInfo clone() {
        return (CellErrorInfo) super.clone();
    }
}

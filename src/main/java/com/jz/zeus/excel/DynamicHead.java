package com.jz.zeus.excel;

import cn.hutool.core.lang.Assert;
import cn.hutool.core.util.StrUtil;
import lombok.Getter;
import lombok.Setter;

/**
 * @author:JZ
 * @date:2021/4/13
 */
@Getter
@Setter
public class DynamicHead {

    private static final Integer DEFAULT_ROW_INDEX = 0;

    private Integer rowIndex;

    private Integer columnIndex;

    private String fieldName;

    private String newName;

    private String appendInfo;

    private DynamicHead() {}

    public static DynamicHead buildAppendInfo(String fieldName, String appendInfo) {
        return build(fieldName, null, appendInfo);
    }

    public static DynamicHead buildAppendInfo(Integer columnIndex, String appendInfo) {
        return build(columnIndex, null, appendInfo);
    }

    public static DynamicHead buildNewName(Integer columnIndex, String newName) {
        return build(columnIndex, newName, null);
    }

    public static DynamicHead buildNewName(String fieldName, String newName) {
        return build(fieldName, newName, null);
    }

    public static DynamicHead build(Integer columnIndex, String newName, String appendInfo) {
        return build(DEFAULT_ROW_INDEX, columnIndex, newName, appendInfo);
    }

    public static DynamicHead build(String fieldName, String newName, String appendInfo) {
        return build(DEFAULT_ROW_INDEX, fieldName, newName, appendInfo);
    }

    public static DynamicHead build(Integer rowIndex, Integer columnIndex, String newName, String appendInfo) {
        if (rowIndex == null) {
            rowIndex = DEFAULT_ROW_INDEX;
        }
        Assert.isTrue(rowIndex >= 0, "rowIndex must be greater than or equal to 0");
        Assert.isTrue(columnIndex != null, "columnIndex must not be null");
        Assert.isTrue(columnIndex >= 0, "columnIndex must be greater than or equal to 0");
        DynamicHead dynamicHead = new DynamicHead();
        dynamicHead.setRowIndex(rowIndex);
        dynamicHead.setColumnIndex(columnIndex);
        dynamicHead.setNewName(newName);
        dynamicHead.setAppendInfo(appendInfo);
        return dynamicHead;
    }

    public static DynamicHead build(Integer rowIndex, String fieldName, String newName, String appendInfo) {
        if (rowIndex == null) {
            rowIndex = DEFAULT_ROW_INDEX;
        }
        Assert.isTrue(rowIndex >= 0, "rowIndex must be greater than or equal to 0");
        Assert.isTrue(StrUtil.isNotBlank(fieldName), "fieldName must not be blank");
        DynamicHead dynamicHead = new DynamicHead();
        dynamicHead.setRowIndex(rowIndex);
        dynamicHead.setFieldName(fieldName);
        dynamicHead.setNewName(newName);
        dynamicHead.setAppendInfo(appendInfo);
        return dynamicHead;
    }

    public String getFinalHeadName(String oldHeadName) {
        if (StrUtil.isBlank(newName)) {
            return StrUtil.concat(true, oldHeadName, appendInfo);
        }
        return StrUtil.concat(true, newName, appendInfo);
    }

}

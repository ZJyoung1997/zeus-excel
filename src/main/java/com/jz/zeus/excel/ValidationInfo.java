package com.jz.zeus.excel;

import cn.hutool.core.collection.CollUtil;
import cn.hutool.core.lang.Assert;
import cn.hutool.core.util.ArrayUtil;
import cn.hutool.core.util.StrUtil;
import com.jz.zeus.excel.util.ClassUtils;
import lombok.Getter;
import lombok.Setter;

import java.util.*;

/**
 * @Author JZ
 * @Date 2021/3/23 18:27
 */
@Getter
@Setter
public class ValidationInfo {

    private static final Integer DEFAULT_ROW_NUM = 10000;

    private static final Integer DEFAULT_COLUMN_NUM = 100;

    private Integer rowIndex;

    private Integer rowNum;

    private Integer columnIndex;

    private Integer columnNum = 100;

    private String fieldName;

    private String headName;

    private List<String> options;

    private ValidationInfo() {}


    public static ValidationInfo buildColumnByIndex(int columnIndex, String... options) {
        return buildColumn(null, null, columnIndex, null, options);
    }

    public static ValidationInfo buildColumnByIndex(int columnIndex, List<String> options) {
        return buildColumn(null, null, columnIndex, null, options);
    }

    public static ValidationInfo buildColumnByIndex(int columnIndex, Integer rowNum, String... options) {
        return buildColumn(null, null, columnIndex, rowNum, options);
    }

    public static ValidationInfo buildColumnByIndex(int columnIndex, Integer rowNum, List<String> options) {
        return buildColumn(null, null, columnIndex, rowNum, options);
    }

    public static ValidationInfo buildColumnByHead(String headName, String... options) {
        return buildColumn(null, headName, null, null, options);
    }
    public static ValidationInfo buildColumnByHead(String headName, List<String> options) {
        return buildColumn(null, headName, null, null, options);
    }

    public static ValidationInfo buildColumnByHead(String headName, Integer rowNum, String... options) {
        return buildColumn(null, headName, null, rowNum, options);
    }

    public static ValidationInfo buildColumnByHead(String headName, Integer rowNum, List<String> options) {
        return buildColumn(null, headName, null, rowNum, options);
    }

    public static ValidationInfo buildColumnByField(String fieldName, String... options) {
        return buildColumn(fieldName, null, null, null, options);
    }

    public static ValidationInfo buildColumnByField(String fieldName, List<String> options) {
        return buildColumn(fieldName, null, null, null, options);
    }

    public static ValidationInfo buildColumnByField(String fieldName, Integer rowNum, String... options) {
        return buildColumn(fieldName, null, null, rowNum, options);
    }

    public static ValidationInfo buildColumnByField(String fieldName, Integer rowNum, List<String> options) {
        return buildColumn(fieldName, null, null, rowNum, options);
    }

    public static ValidationInfo buildColumnByField(Class<?> clazz, String fieldName, Integer rowNum, String... options) {
        if (ClassUtils.getFieldInfoByFieldName(clazz, fieldName).isPresent()) {
            return buildColumn(fieldName, null, null, rowNum, options);
        }
        return null;
    }

    public static ValidationInfo buildColumn(String fieldName, String headName, Integer columnIndex, Integer rowNum, String... options) {
        if (ArrayUtil.isEmpty(options)) {
            return buildColumn(fieldName, headName, columnIndex, rowNum, new ArrayList<>(0));
        }
        return buildColumn(fieldName, headName, columnIndex, rowNum, CollUtil.toList(options));
    }

    public static ValidationInfo buildColumn(String fieldName, String headName, Integer columnIndex, Integer rowNum, List<String> options) {
        Assert.isFalse(columnIndex == null && StrUtil.isBlank(fieldName) && StrUtil.isBlank(headName),
                "ColumnIndex and FieldName and HeadName not all empty");
        if (rowNum != null) {
            Assert.isTrue(rowNum >= 0, "RowNum has to be greater than or equal to 0");
        }
        if (columnIndex != null) {
            Assert.isTrue(columnIndex >= 0, "ColumnIndex has to be greater than or equal to 0");
        }
        ValidationInfo validationInfo = new ValidationInfo();
        validationInfo.setColumnIndex(columnIndex);
        validationInfo.setFieldName(fieldName);
        validationInfo.setHeadName(headName);
        validationInfo.setRowNum(rowNum == null ? DEFAULT_ROW_NUM : rowNum);
        validationInfo.setOptions(options);
        return validationInfo;
    }

    /**
     * 构建具有相同选项的多列下拉框
     * @param names                 表头 或 属性名
     * @param isField               true names 表示属性名、false names 表示表头
     * @param options               选项内容
     */
    public static List<ValidationInfo> buildCommonOption(String[] names, boolean isField, String... options) {
        if (ArrayUtil.isEmpty(names)) {
            return new ArrayList<>(0);
        }
        List<ValidationInfo> validationInfoList = new ArrayList<>();
        for(String name : names) {
            if (isField) {
                validationInfoList.add(ValidationInfo.buildColumnByField(name, options));
            } else {
                validationInfoList.add(ValidationInfo.buildColumnByHead(name, options));
            }
        }
        return validationInfoList;
    }

    /**
     * 构建具有相同选项的多列下拉框
     * @param columnIndexs          列索引
     * @param options               选项内容
     */
    public static List<ValidationInfo> buildCommonOption(Integer[] columnIndexs, String... options) {
        if (ArrayUtil.isEmpty(columnIndexs) || ArrayUtil.isEmpty(options)) {
            return new ArrayList<>(0);
        }
        List<ValidationInfo> validationInfoList = new ArrayList<>(columnIndexs.length);
        for(int i = 0; i < columnIndexs.length; ++i) {
            validationInfoList.add(ValidationInfo.buildColumnByIndex(columnIndexs[i], options));
        }
        return validationInfoList;
    }

    /**
     * 构建某一行的下拉框
     */
    public static ValidationInfo bulidRow(Integer rowIndex, String... options) {
        return buildRow(rowIndex, null, options);
    }

    /**
     * 构建某一行的下拉框
     */
    public static ValidationInfo buildRow(Integer rowIndex, Integer columnNum, String... options) {
        Assert.isTrue(rowIndex != null, "RowIndex can't be null");
        ValidationInfo info = new ValidationInfo();
        info.setRowIndex(rowIndex);
        info.setColumnNum(columnNum == null ? DEFAULT_COLUMN_NUM : columnNum);
        info.setOptions(options == null ? Collections.emptyList() : Arrays.asList(options));
        return info;
    }

    /**
     * 构建某个单元格的下拉框
     */
    public static ValidationInfo buildPrecise(Integer rowIndex, Integer columnIndex, String... options) {
        Assert.isTrue(rowIndex != null && columnIndex != null, "RowIndex and ColumnIndex can't both be null");
        ValidationInfo info = new ValidationInfo();
        info.setRowIndex(rowIndex);
        info.setColumnIndex(columnIndex);
        info.setOptions(options == null ? Collections.emptyList() : Arrays.asList(options));
        return info;
    }

    /**
     * 构建某个单元格的下拉框
     */
    public static ValidationInfo buildPrecise(Integer rowIndex, String headName, String... options) {
        Assert.isTrue(rowIndex != null && StrUtil.isNotBlank(headName), "RowIndex and HeadName can't both be empty");
        ValidationInfo info = new ValidationInfo();
        info.setRowIndex(rowIndex);
        info.setHeadName(headName);
        info.setOptions(options == null ? Collections.emptyList() : Arrays.asList(options));
        return info;
    }

    @Override
    public boolean equals(Object o) {
        if (this == o) {
            return true;
        }
        if (!(o instanceof ValidationInfo)) {
            return false;
        }
        ValidationInfo that = (ValidationInfo) o;
        if (Objects.equals(rowIndex, that.rowIndex)) {
            if (rowIndex != null && Objects.equals(that.columnIndex, columnIndex)) {
                return true;
            } else if (StrUtil.isNotBlank(fieldName) && Objects.equals(that.fieldName, fieldName)) {
                return true;
            } else if (StrUtil.isNotBlank(headName) && Objects.equals(that.headName, headName)) {
                return true;
            }
        }
        return false;
    }

    @Override
    public int hashCode() {
        return Objects.hash(rowIndex, fieldName, rowNum, columnIndex, columnNum, headName, options);
    }
}

package com.jz.zeus.excel;

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
public class DropDownBoxInfo {

    private static final Integer DEFAULT_ROW_NUM = 10000;

    private static final Integer DEFAULT_COLUMN_NUM = 100;

    private Integer rowIndex;

    private Integer rowNum;

    private Integer columnIndex;

    private Integer columnNum = 100;

    private String headName;

    private List<String> options;

    private DropDownBoxInfo() {}


    public static DropDownBoxInfo buildColumn(String headName, String... options) {
        return buildColumn(headName, null, options);
    }

    public static DropDownBoxInfo buildColumn(String headName, Integer rowNum, String... options) {
        return buildColumn(headName, rowNum, options == null ? Collections.emptyList() : Arrays.asList(options));
    }

    public static DropDownBoxInfo buildColumn(String headName, List<String> options) {
        return buildColumn(headName, null, options);
    }

    public static DropDownBoxInfo buildColumn(String headName, Integer rowNum, List<String> options) {
        Assert.isTrue(StrUtil.isNotBlank(headName), "HeadName can't be empty");
        DropDownBoxInfo dropDownBoxInfo = new DropDownBoxInfo();
        dropDownBoxInfo.setHeadName(headName);
        dropDownBoxInfo.setRowNum(rowNum == null ? DEFAULT_ROW_NUM : rowNum);
        dropDownBoxInfo.setOptions(options);
        return dropDownBoxInfo;
    }

    public static DropDownBoxInfo buildColumn(Integer columnIndex, String... options) {
        return buildColumn(columnIndex, null, options);
    }

    public static DropDownBoxInfo buildColumn(Integer columnIndex, Integer rowNum, String... options) {
        return buildColumn(columnIndex, rowNum, options == null ? Collections.emptyList() : Arrays.asList(options));
    }

    /**
     * 构建某一列的下拉框
     * @param columnIndex     列索引
     * @param options         选项内容
     */
    public static DropDownBoxInfo buildColumn(Integer columnIndex, List<String> options) {
        return buildColumn(columnIndex, null, options);
    }

    /**
     * 构建某一列的下拉框
     * @param columnIndex     列索引
     * @param rowNum          需要构建下拉框的行数
     * @param options         选项内容
     */
    public static DropDownBoxInfo buildColumn(Integer columnIndex, Integer rowNum, List<String> options) {
        Assert.isTrue(columnIndex != null, "ColumnIndex can't be null");
        DropDownBoxInfo dropDownBoxInfo = new DropDownBoxInfo();
        dropDownBoxInfo.setColumnIndex(columnIndex);
        dropDownBoxInfo.setRowNum(rowNum == null ? DEFAULT_ROW_NUM : rowNum);
        dropDownBoxInfo.setOptions(options);
        return dropDownBoxInfo;
    }

    public static DropDownBoxInfo buildColumn(Class<?> clazz, String fieldName, List<String> options) {
        return buildColumn(clazz, fieldName, null, options);
    }

    public static DropDownBoxInfo buildColumn(Class<?> clazz, String fieldName, Integer rowNum, List<String> options) {
        Optional<FieldInfo> fieldInfoOptional = ClassUtils.getFieldInfoByFieldName(clazz, fieldName);
        if (fieldInfoOptional.isPresent()) {
            FieldInfo fieldInfo = fieldInfoOptional.get();
            if (fieldInfo.getHeadColumnIndex() != null) {
                return buildColumn(fieldInfo.getHeadColumnIndex(), rowNum, options);
            } else if (StrUtil.isNotBlank(fieldInfo.getHeadName())) {
                return buildColumn(fieldInfo.getHeadName(), rowNum, options);
            }
        }
        return null;
    }

    /**
     * 构建具有相同选项的多列下拉框
     * @param headNames             表头
     * @param options               选项内容
     */
    public static List<DropDownBoxInfo> buildCommonOption(String[] headNames, String... options) {
        if (ArrayUtil.isEmpty(headNames) || ArrayUtil.isEmpty(options)) {
            return new ArrayList<>(0);
        }
        List<DropDownBoxInfo> dropDownBoxInfoList = new ArrayList<>();
        List<String> optionList = Arrays.asList(options);
        for(int i = 0; i < headNames.length; ++i) {
            dropDownBoxInfoList.add(DropDownBoxInfo.buildColumn(headNames[i], optionList));
        }
        return dropDownBoxInfoList;
    }

    /**
     * 构建具有相同选项的多列下拉框
     * @param columnIndexs          列索引
     * @param options               选项内容
     */
    public static List<DropDownBoxInfo> buildCommonOption(Integer[] columnIndexs, String... options) {
        if (ArrayUtil.isEmpty(columnIndexs) || ArrayUtil.isEmpty(options)) {
            return new ArrayList<>(0);
        }
        List<DropDownBoxInfo> dropDownBoxInfoList = new ArrayList<>();
        List<String> optionList = Arrays.asList(options);
        for(int i = 0; i < columnIndexs.length; ++i) {
            dropDownBoxInfoList.add(DropDownBoxInfo.buildColumn(columnIndexs[i], optionList));
        }
        return dropDownBoxInfoList;
    }

    /**
     * 构建某一行的下拉框
     */
    public static DropDownBoxInfo bulidRow(Integer rowIndex, String... options) {
        return buildRow(rowIndex, null, options);
    }

    /**
     * 构建某一行的下拉框
     */
    public static DropDownBoxInfo buildRow(Integer rowIndex, Integer columnNum, String... options) {
        Assert.isTrue(rowIndex != null, "RowIndex can't be null");
        DropDownBoxInfo info = new DropDownBoxInfo();
        info.setRowIndex(rowIndex);
        info.setColumnNum(columnNum == null ? DEFAULT_COLUMN_NUM : columnNum);
        info.setOptions(options == null ? Collections.emptyList() : Arrays.asList(options));
        return info;
    }

    /**
     * 构建某个单元格的下拉框
     */
    public static DropDownBoxInfo buildPrecise(Integer rowIndex, Integer columnIndex, String... options) {
        Assert.isTrue(rowIndex != null && columnIndex != null, "RowIndex and ColumnIndex can't both be null");
        DropDownBoxInfo info = new DropDownBoxInfo();
        info.setRowIndex(rowIndex);
        info.setColumnIndex(columnIndex);
        info.setOptions(options == null ? Collections.emptyList() : Arrays.asList(options));
        return info;
    }

    /**
     * 构建某个单元格的下拉框
     */
    public static DropDownBoxInfo buildPrecise(Integer rowIndex, String headName, String... options) {
        Assert.isTrue(rowIndex != null && StrUtil.isNotBlank(headName), "RowIndex and HeadName can't both be empty");
        DropDownBoxInfo info = new DropDownBoxInfo();
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
        if (!(o instanceof DropDownBoxInfo)) {
            return false;
        }
        DropDownBoxInfo that = (DropDownBoxInfo) o;
        return (Objects.equals(rowIndex, that.rowIndex) && Objects.equals(columnIndex, that.columnIndex)) ||
                (Objects.equals(rowIndex, that.rowIndex) && Objects.equals(headName, that.headName));
    }

    @Override
    public int hashCode() {
        return Objects.hash(rowIndex, rowNum, columnIndex, columnNum, headName, options);
    }
}

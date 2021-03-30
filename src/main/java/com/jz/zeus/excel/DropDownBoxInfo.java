package com.jz.zeus.excel;

import cn.hutool.core.lang.Assert;
import cn.hutool.core.util.StrUtil;
import lombok.Getter;
import lombok.Setter;

import java.util.Arrays;
import java.util.Collections;
import java.util.List;

/**
 * @Author JZ
 * @Date 2021/3/23 18:27
 */
@Getter
@Setter
public class DropDownBoxInfo {

    private static final Integer DEFAULT_ROW_NUM = 100;

    private static final Integer DEFAULT_COLUMN_NUM = 100;

    private Integer rowIndex;

    private Integer rowNum;

    private Integer columnIndex;

    private Integer columnNum = 100;

    private String headName;

    private List<String> options;

    private DropDownBoxInfo() {}

    public DropDownBoxInfo(String headName, String... options) {
        this(headName, null, options);
    }


    public DropDownBoxInfo(String headName, List<String> options) {
        this(headName, null, options);
    }

    public DropDownBoxInfo(String headName, Integer rowNum, String... options) {
        Assert.isTrue(StrUtil.isNotBlank(headName), "HeadName can't be empty");
        this.headName = headName;
        this.rowNum = rowNum == null ? DEFAULT_ROW_NUM : rowNum;
        this.options = options == null ? Collections.emptyList() : Arrays.asList(options);
    }

    public DropDownBoxInfo(String headName, Integer rowNum, List<String> options) {
        Assert.isTrue(StrUtil.isNotBlank(headName), "HeadName can't be empty");
        this.headName = headName;
        this.rowNum = rowNum == null ? DEFAULT_ROW_NUM : rowNum;
        this.options = options;
    }

    public DropDownBoxInfo(Integer columnIndex, String... options) {
        this(columnIndex, null, options);
    }

    public DropDownBoxInfo(Integer columnIndex, List<String> options) {
        this(columnIndex, null, options);
    }

    public DropDownBoxInfo(Integer columnIndex, Integer rowNum, String... options) {
        Assert.isTrue(columnIndex != null, "ColumnIndex can't be null");
        this.columnIndex = columnIndex;
        this.rowNum = rowNum == null ? DEFAULT_ROW_NUM : rowNum;
        this.options = options == null ? Collections.emptyList() : Arrays.asList(options);
    }

    public DropDownBoxInfo(Integer columnIndex, Integer rowNum, List<String> options) {
        Assert.isTrue(columnIndex != null, "ColumnIndex can't be null");
        this.columnIndex = columnIndex;
        this.rowNum = rowNum == null ? DEFAULT_ROW_NUM : rowNum;
        this.options = options;
    }

    public static DropDownBoxInfo getRowDropDownBoxInfo(Integer rowIndex, String... options) {
        return getRowDropDownBoxInfo(rowIndex, null, options);
    }

    public static DropDownBoxInfo getRowDropDownBoxInfo(Integer rowIndex, Integer columnNum, String... options) {
        Assert.isTrue(rowIndex != null, "RowIndex can't be null");
        DropDownBoxInfo info = new DropDownBoxInfo();
        info.setRowIndex(rowIndex);
        info.setColumnNum(columnNum == null ? DEFAULT_COLUMN_NUM : columnNum);
        info.setOptions(options == null ? Collections.emptyList() : Arrays.asList(options));
        return info;
    }

    public static DropDownBoxInfo getInstance(Integer rowIndex, Integer columnIndex, String... options) {
        Assert.isTrue(rowIndex != null && columnIndex != null, "RowIndex and ColumnIndex can't both be null");
        DropDownBoxInfo info = new DropDownBoxInfo();
        info.setRowIndex(rowIndex);
        info.setColumnIndex(columnIndex);
        info.setOptions(options == null ? Collections.emptyList() : Arrays.asList(options));
        return info;
    }

    public static DropDownBoxInfo getInstance(Integer rowIndex, String headName, String... options) {
        Assert.isTrue(rowIndex != null && StrUtil.isNotBlank(headName), "RowIndex and HeadName can't both be empty");
        DropDownBoxInfo info = new DropDownBoxInfo();
        info.setRowIndex(rowIndex);
        info.setHeadName(headName);
        info.setOptions(options == null ? Collections.emptyList() : Arrays.asList(options));
        return info;
    }

}

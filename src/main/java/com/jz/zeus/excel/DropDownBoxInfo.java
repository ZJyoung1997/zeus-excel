package com.jz.zeus.excel;

import cn.hutool.core.lang.Assert;
import cn.hutool.core.util.StrUtil;
import lombok.Data;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

import java.util.Arrays;
import java.util.List;

/**
 * @Author JZ
 * @Date 2021/3/23 18:27
 */
@Getter
@Setter
@NoArgsConstructor
public class DropDownBoxInfo {

    private Integer rowIndex;

    private Integer rowNum;

    private Integer columnIndex;

    private Integer columnNum;

    private String headName;

    private List<String> options;

    public DropDownBoxInfo(String headName, String... options) {
        this(headName, null, options);
    }

    public DropDownBoxInfo(String headName, Integer rowNum, String... options) {
        Assert.isTrue(StrUtil.isNotBlank(headName), "HeadName can't be empty");
        this.headName = headName;
        this.rowNum = rowNum;
        this.options = Arrays.asList(options);
    }

    public DropDownBoxInfo(Integer columnIndex, String... options) {
        this(columnIndex, null, options);
    }

    public DropDownBoxInfo(Integer columnIndex, Integer rowNum, String... options) {
        Assert.isTrue(columnIndex != null, "ColumnIndex can't be null");
        this.headName = headName;
        this.rowNum = rowNum;
        this.options = Arrays.asList(options);
    }

}

package com.jz.zeus.excel;

import cn.hutool.core.collection.CollUtil;
import cn.hutool.core.lang.Assert;
import cn.hutool.core.util.StrUtil;
import lombok.Data;

import java.util.ArrayList;
import java.util.Collection;

/**
 * @Author JZ
 * @Date 2021/3/22 17:56
 */
@Data
public class CellErrorInfo {

    private Integer rowIndex;

    private Integer columnIndex;

    private String headName;

    private Collection<String> errorMsg = new ArrayList<>();

    public CellErrorInfo(int rowIndex, String headName) {
        this(rowIndex, null, headName);
    }

    public CellErrorInfo(int rowIndex, int columnIndex) {
        this(rowIndex, Integer.valueOf(columnIndex), null);
    }

    public CellErrorInfo(int rowIndex, Integer columnIndex, String headName) {
        Assert.isFalse(columnIndex == null && StrUtil.isBlank(headName),
                "ColumnIndex and HeadName can't both be empty");
        Assert.isTrue(rowIndex >= 0, "RowIndex has to be greater than or equal to 0");
        this.rowIndex = Integer.valueOf(rowIndex);
        if (columnIndex != null) {
            Assert.isTrue(columnIndex >= 0, "ColumnIndex has to be greater than or equal to 0");
            this.columnIndex = columnIndex;
        }
        if (StrUtil.isNotBlank(headName)) {
            this.headName = headName;
        }
    }

    public void addErrorMsg(String msg) {
        if (StrUtil.isNotBlank(msg)) {
            errorMsg.add(msg);
        }
    }

    public boolean hasError() {
        return CollUtil.isNotEmpty(errorMsg);
    }

}

package com.jz.zeus.excel.write.builder;

import cn.hutool.core.lang.Assert;
import com.alibaba.excel.write.metadata.WriteSheet;
import com.jz.zeus.excel.context.ExcelContext;
import lombok.Getter;

import java.util.List;

/**
 * @Author JZ
 * @Date 2021/4/23 15:35
 */
public class ZeusWriteSheet extends WriteSheet {

    @Getter
    private ExcelContext excelContext;

    public ZeusWriteSheet(ExcelContext excelContext) {
        Assert.notNull(excelContext, "ExcelContext must not be null");
        this.excelContext = excelContext;
    }

    public void updateDatas(List data) {
        excelContext.setSheetData(data);
    }

    public void addDatas(List data) {
        excelContext.addSheetData(data);
    }

}

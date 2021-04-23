package com.jz.zeus.excel.write.builder;

import cn.hutool.core.lang.Assert;
import com.alibaba.excel.write.metadata.WriteSheet;
import com.jz.zeus.excel.context.ExcelContext;
import lombok.Getter;
import lombok.Setter;

import java.util.List;

/**
 * @Author JZ
 * @Date 2021/4/23 15:35
 */
public class ZeusWriterSheet extends WriteSheet {

    @Getter
    @Setter
    private ExcelContext excelContext;

    public ZeusWriterSheet() {
        this.excelContext = new ExcelContext();
    }

    public ZeusWriterSheet(ExcelContext excelContext) {
        Assert.notNull(excelContext, "ExcelContext must not be null");
        this.excelContext = excelContext;
    }

    public void updateDatas(List data) {
        excelContext.setSheetData(data);
    }

}

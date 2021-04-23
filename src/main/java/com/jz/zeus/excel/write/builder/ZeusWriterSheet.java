package com.jz.zeus.excel.write.builder;

import com.alibaba.excel.write.metadata.WriteSheet;
import com.jz.zeus.excel.context.ExcelContext;
import lombok.Getter;
import lombok.Setter;

/**
 * @Author JZ
 * @Date 2021/4/23 15:35
 */
public class ZeusWriterSheet extends WriteSheet {

    @Getter
    @Setter
    private ExcelContext excelContext;

}

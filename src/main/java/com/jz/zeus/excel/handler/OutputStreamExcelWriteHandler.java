package com.jz.zeus.excel.handler;

import cn.hutool.core.collection.CollectionUtil;
import cn.hutool.core.util.StrUtil;
import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.write.builder.ExcelWriterBuilder;
import com.alibaba.excel.write.builder.ExcelWriterSheetBuilder;
import com.alibaba.excel.write.metadata.WriteSheet;
import com.jz.zeus.excel.ExcelInfo;
import com.jz.zeus.excel.SheetInfo;

import java.io.OutputStream;

/**
 * @author:JZ
 * @date:2021/3/21
 */
public class OutputStreamExcelWriteHandler implements ExcelWriteHandler<OutputStream> {

    @Override
    public boolean check(ExcelInfo excelInfo) {
        if (excelInfo == null) {
            throw new IllegalArgumentException("ExcelInfo cannot be null");
        }
        if (StrUtil.isBlank(excelInfo.getName())) {
            throw new IllegalArgumentException("ExcelInfo.name cannot be blank string");
        }
        return true;
    }

    @Override
    public void execute(ExcelInfo excelInfo, OutputStream outputStream) {
        check(excelInfo);
        write(excelInfo, outputStream);
    }

    @Override
    public void write(ExcelInfo excelInfo, OutputStream outputStream) {
        ExcelWriter excelWriter = getExcelWriter(excelInfo, outputStream);
        if (CollectionUtil.isEmpty(excelInfo.getSheetInfos())) {
            return;
        }
        excelInfo.getSheetInfos().forEach(sheetInfo -> {
            WriteSheet writeSheet = getWriteSheet(sheetInfo);
            excelWriter.write(sheetInfo.getDataList(), writeSheet);
        });
        excelWriter.finish();
    }

    public ExcelWriter getExcelWriter(ExcelInfo excelInfo, OutputStream outputStream) {
        ExcelWriterBuilder writerBuilder = EasyExcel.write(outputStream)
                .autoCloseStream(true)
                .excelType(excelInfo.getSuffix())
                .inMemory(excelInfo.isInMemory());

        if (StrUtil.isNotBlank(excelInfo.getPassword())) {
            writerBuilder.password(excelInfo.getPassword());
        }

        return writerBuilder.build();
    }

    public WriteSheet getWriteSheet(SheetInfo sheetInfo) {
        ExcelWriterSheetBuilder sheetBuilder = EasyExcel
                .writerSheet(sheetInfo.getIndex(), sheetInfo.getName());

        if (CollectionUtil.isNotEmpty(sheetInfo.getDataList())) {
            sheetBuilder.head(sheetInfo.getDataList().get(0).getClass());
        }

        if (CollectionUtil.isNotEmpty(sheetInfo.getIncludeHeaderNames())) {
            sheetBuilder.includeColumnFiledNames(sheetInfo.getIncludeHeaderNames());
        }

        if (CollectionUtil.isNotEmpty(sheetInfo.getExcludeHeaderNames())) {
            sheetBuilder.excludeColumnFiledNames(sheetInfo.getExcludeHeaderNames());
        }
        return sheetBuilder.build();
    }

}

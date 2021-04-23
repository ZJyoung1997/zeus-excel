package com.jz.zeus.excel.write.builder;

import cn.hutool.core.lang.Assert;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.context.WriteContext;
import com.alibaba.excel.write.metadata.WriteTable;
import lombok.Getter;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.List;

/**
 * @Author JZ
 * @Date 2021/4/23 17:21
 */
public class ZeusExcelWriter {

    private static final Logger LOGGER = LoggerFactory.getLogger(ZeusExcelWriter.class);

    @Getter
    private ExcelWriter excelWriter;

    public ZeusExcelWriter(ExcelWriter excelWriter) {
        Assert.notNull(excelWriter, "ExcelWriter must not be null");
        this.excelWriter = excelWriter;
    }

    /**
     * Write data to a sheet
     *
     * @param data
     *            Data to be written
     * @param writeSheet
     *            Write to this sheet
     * @return this current writer
     */
    public ZeusExcelWriter write(List data, ZeusWriterSheet writeSheet) {
        return write(data, writeSheet, null);
    }

    /**
     * Write value to a sheet
     *
     * @param data
     *            Data to be written
     * @param writeSheet
     *            Write to this sheet
     * @param writeTable
     *            Write to this table
     * @return this
     */
    public ZeusExcelWriter write(List data, ZeusWriterSheet writeSheet, WriteTable writeTable) {
        writeSheet.updateDatas(data);
        excelWriter.write(data, writeSheet, writeTable);
        return this;
    }

    /**
     * Close IO
     */
    public void finish() {
        excelWriter.finish();
    }

    /**
     * Prevents calls to {@link #finish} from freeing the cache
     *
     */
    @Override
    protected void finalize() {
        try {
            finish();
        } catch (Throwable e) {
            LOGGER.warn("Destroy object failed", e);
        }
    }

    /**
     * The context of the entire writing process
     *
     * @return
     */
    public WriteContext writeContext() {
        return excelWriter.writeContext();
    }

}

package com.jz.zeus.excel.write.builder;

import cn.hutool.core.lang.Assert;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.context.WriteContext;
import com.alibaba.excel.write.metadata.WriteTable;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * @Author JZ
 * @Date 2021/4/23 17:21
 */
public class ZeusExcelWriter {

    private static final Logger LOGGER = LoggerFactory.getLogger(ZeusExcelWriter.class);

    private Map<ZeusWriteSheet, Boolean> writeSheetMap = new HashMap<>();

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
    public ZeusExcelWriter write(List data, ZeusWriteSheet writeSheet) {
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
    public ZeusExcelWriter write(List data, ZeusWriteSheet writeSheet, WriteTable writeTable) {
        Assert.notNull(writeSheet, "Sheet argument cannot be null");
        if (writeSheetMap.get(writeSheet) == null) {
            writeSheet.updateDatas(data);
        } else {
            writeSheet.addDatas(data);
        }
        excelWriter.write(data, writeSheet, writeTable);
        writeSheetMap.put(writeSheet, Boolean.TRUE);
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

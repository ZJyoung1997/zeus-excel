package com.jz.zeus.excel.handler;

import com.jz.zeus.excel.ExcelInfo;

/**
 * @author:JZ
 * @date:2021/3/21
 */
public interface ExcelWriteHandler<T> {

    /**
     * 写入Excel前校验
     * @param excelInfo
     * @return
     */
    boolean check(ExcelInfo excelInfo);

    /**
     * 写入Excel入口方法，调用 check 和 write
     * @param excelInfo
     * @param t
     */
    void execute(ExcelInfo excelInfo, T t);

    /**
     * 实际写入Excel方法
     * @param excelInfo
     * @param t
     */
    void write(ExcelInfo excelInfo, T t);

}

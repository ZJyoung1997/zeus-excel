package com.jz.zeus.excel.test;

import com.alibaba.excel.EasyExcel;

/**
 * @Author JZ
 * @Date 2021/3/22 10:17
 */
public class ExcelTest {

    public static void main(String[] args) {
        read();
    }

    public static void read() {
        String fileName = "C:\\Users\\User\\Desktop\\254.xlsx";
        EasyExcel.read(fileName, DemoData.class, new DemoExcelReadListener(1))
                .sheet("模板").doRead();
    }

    public static void write() {
    }

}

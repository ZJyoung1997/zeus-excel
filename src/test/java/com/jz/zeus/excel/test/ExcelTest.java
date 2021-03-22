package com.jz.zeus.excel.test;

import com.alibaba.excel.EasyExcel;
import com.jz.zeus.excel.CellErrorInfo;
import com.jz.zeus.excel.write.handler.CellErrorInfoCommentHandler;

import java.util.ArrayList;
import java.util.List;

/**
 * @Author JZ
 * @Date 2021/3/22 10:17
 */
public class ExcelTest {

    public static void main(String[] args) {
//        read();
        write();
    }

    public static void read() {
        String fileName = "C:\\Users\\Administrator\\Desktop\\254.xlsx";
        EasyExcel.read(fileName, DemoData.class, new DemoExcelReadListener(1))
                .sheet("模板").doRead();
    }

    public static void write() {
        String fileName = "C:\\Users\\Administrator\\Desktop\\254.xlsx";
        EasyExcel.write(fileName, DemoData.class).sheet("模板")
                .registerWriteHandler(new CellErrorInfoCommentHandler((List<CellErrorInfo>) null))
                .doWrite(getDataList());
    }

    private static List<DemoData> getDataList() {
        List<DemoData> dataList = new ArrayList<>();
        for (int i = 0; i < 10; i++) {
            DemoData demoData = new DemoData();
            demoData.setMateCode("mateCode" + i);
            demoData.setDest("dest" + i);
            demoData.setSrc("src" + i);
            demoData.setFunc("func" + i);
            dataList.add(demoData);
        }
        return dataList;
    }

}

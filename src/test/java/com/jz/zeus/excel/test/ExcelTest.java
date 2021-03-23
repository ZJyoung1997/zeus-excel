package com.jz.zeus.excel.test;

import com.alibaba.excel.EasyExcel;
import com.jz.zeus.excel.CellErrorInfo;
import com.jz.zeus.excel.write.handler.CellErrorInfoCommentHandler;
import com.jz.zeus.excel.write.handler.DefaultHeadStyleHandler;
import com.jz.zeus.excel.write.handler.DropDownBoxHandler;

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
//        String fileName = "C:\\Users\\Administrator\\Desktop\\254.xlsx";
        String fileName = "C:\\Users\\User\\Desktop\\254.xlsx";
        EasyExcel.read(fileName, DemoData.class, new DemoExcelReadListener(1))
                .sheet("模板").doRead();
    }

    public static void write() {
        String fileName = "C:\\Users\\Administrator\\Desktop\\254.xlsx";
//        String fileName = "C:\\Users\\User\\Desktop\\254.xlsx";
        List<CellErrorInfo> cellErrorInfoList = new ArrayList<>();
        cellErrorInfoList.add(new CellErrorInfo(1, 1, "格式错误"));
        cellErrorInfoList.add(new CellErrorInfo(4, "媒体CODE", "关系错误"));
        cellErrorInfoList.add(new CellErrorInfo(2, "FUNC", "格式错误")
                .addErrorMsg("数值放假看电视了积分卡积分错误"));

        EasyExcel.write(fileName).sheet("模板").head(DemoData.class)
                .registerWriteHandler(new DefaultHeadStyleHandler())
                .registerWriteHandler(new CellErrorInfoCommentHandler(cellErrorInfoList))
                .registerWriteHandler(new DropDownBoxHandler())
//                .excludeColumnFiledNames(Arrays.asList("dest"))
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

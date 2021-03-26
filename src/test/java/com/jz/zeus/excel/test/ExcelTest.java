package com.jz.zeus.excel.test;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.read.listener.ReadListener;
import com.jz.zeus.excel.CellErrorInfo;
import com.jz.zeus.excel.DropDownBoxInfo;
import com.jz.zeus.excel.read.listener.AbstractExcelReadListener;
import com.jz.zeus.excel.util.ExcelUtils;
import com.jz.zeus.excel.write.handler.DefaultHeadStyleHandler;
import com.jz.zeus.excel.write.handler.DropDownBoxSheetHandler;
import com.jz.zeus.excel.write.handler.ErrorInfoCommentHandler;
import lombok.SneakyThrows;

import java.io.*;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.List;

/**
 * @Author JZ
 * @Date 2021/3/22 10:17
 */
public class ExcelTest {

    public static void main(String[] args) throws IOException {
//        String path = "C:\\Users\\Administrator\\Desktop\\254.xlsx";
        String path = "C:\\Users\\User\\Desktop\\254.xlsx";

        ExcelUtils.createTemplate(new FileOutputStream(path), "模板", DemoData.class, getDropDownBoxInfo(), Arrays.asList("src"));

//        write(new FileOutputStream(path), null);

        AbstractExcelReadListener readListener = new DemoExcelReadListener(5);
//        ExcelUtils.readAndWriteErrorMsg(readListener, path, "模板", DemoData.class);

//        read(new FileInputStream(path), readListener);
//        ExcelUtils.addErrorInfo(path, path, "模板", readListener.getErrorInfoList());

        System.out.println("end");
    }

    public static void read(InputStream inputStream, ReadListener readListener) {
        EasyExcel.read(inputStream)
                .sheet("模板")
                .head(DemoData.class)
                .registerReadListener(readListener)
                .doRead();
    }

    @SneakyThrows
    public static void write(OutputStream outputStream, List<CellErrorInfo> cellErrorInfoList) {
        EasyExcel.write(outputStream)
                .sheet("模板").head(DemoData.class)
                .registerWriteHandler(new DefaultHeadStyleHandler())
                .registerWriteHandler(new DropDownBoxSheetHandler(getDropDownBoxInfo()))
//                .excludeColumnFiledNames(Arrays.asList("dest"))
//                .doWrite(Collections.emptyList());
                .doWrite(getDataList());
        outputStream.close();
    }

    @SneakyThrows
    public static void fill(OutputStream outputStream, InputStream inputStream, List<CellErrorInfo> cellErrorInfoList) {

        EasyExcel.write(outputStream)
                .withTemplate(inputStream)
                .sheet("模板")
                .registerWriteHandler(new ErrorInfoCommentHandler(cellErrorInfoList))
                .doWrite(Collections.emptyList());
        inputStream.close();
        outputStream.close();
    }

    public static List<DropDownBoxInfo> getDropDownBoxInfo() {
        List<DropDownBoxInfo> dropDownBoxInfoList = new ArrayList<>();
        dropDownBoxInfoList.add(new DropDownBoxInfo("SRC", "是", "否"));
        dropDownBoxInfoList.add(new DropDownBoxInfo(1,"可以", "不可以"));
        dropDownBoxInfoList.add(DropDownBoxInfo.getRowDropDownBoxInfo(2, "中", "不中"));
        dropDownBoxInfoList.add(DropDownBoxInfo.getInstance(3, "媒体CODE", "不中"));
        return dropDownBoxInfoList;
    }

    public static List<CellErrorInfo> getCellErrorInfo() {
        List<CellErrorInfo> cellErrorInfoList = new ArrayList<>();
        cellErrorInfoList.add(new CellErrorInfo(1, 1, "格式错误"));
        cellErrorInfoList.add(new CellErrorInfo(4, "媒体CODE", "关系错误"));
        cellErrorInfoList.add(new CellErrorInfo(2, "FUNC", "格式错误")
                .addErrorMsg("数值放假看电视了积分卡积分错误"));
        return cellErrorInfoList;
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

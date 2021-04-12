package com.jz.zeus.excel.test;

import com.alibaba.excel.EasyExcel;
import com.jz.zeus.excel.CellErrorInfo;
import com.jz.zeus.excel.ValidationInfo;
import com.jz.zeus.excel.read.listener.ExcelReadListener;
import com.jz.zeus.excel.test.data.DemoData;
import com.jz.zeus.excel.test.listener.DemoExcelReadListener;
import com.jz.zeus.excel.util.ExcelUtils;
import com.jz.zeus.excel.write.handler.ErrorInfoCommentHandler;
import com.jz.zeus.excel.write.handler.ExtendColumnHandler;
import com.jz.zeus.excel.write.handler.HeadStyleHandler;
import com.jz.zeus.excel.write.property.CellStyleProperty;
import lombok.SneakyThrows;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.*;

/**
 * @Author JZ
 * @Date 2021/3/22 10:17
 */
public class ExcelTest {

    public static void main(String[] args) throws IOException {
//        String path = "C:\\Users\\Administrator\\Desktop\\254.xlsx";
        String path = "C:\\Users\\User\\Desktop\\254.xlsx";
        long startTime = System.currentTimeMillis();
        CellStyleProperty styleProperty = CellStyleProperty.getDefaultHeadProperty();
        styleProperty.setFillPatternType(FillPatternType.SOLID_FOREGROUND);
        styleProperty.setFillForegroundColor(IndexedColors.RED.index);
        List<CellStyleProperty> list = new ArrayList<>();
        list.add(styleProperty);

        List<ValidationInfo> validationInfos = new ArrayList<ValidationInfo>() {{
            add(ValidationInfo.buildColumn("ID", "是", "否"));
        }};
//        ExcelUtils.createTemplate(new FileOutputStream(path), "模板", Arrays.asList("媒体发发发CODE", "解不不不不不决"), new HeadStyleHandler(styleProperty), getDropDownBoxInfo());
//        ExcelUtils.createTemplate(new FileOutputStream(path), "模板", DemoData.class, new HeadStyleHandler(), null, null);
//        ExcelUtils.createTemplate(path, "模板", Arrays.asList("jj", "jkfk"), new HeadStyleHandler(list), null);

//        ExcelUtils.write(path, "模板", Arrays.asList("字符串", "数字", "dest"), getDataList1(getHead()), null, null);
//        ExcelUtils.write(path, "模板", DemoData.class, getDataList(), null, null, null);
        write(new FileOutputStream(path), getCellErrorInfo());

        ExcelReadListener readListener = new DemoExcelReadListener(5);
//        ExcelUtils.readAndWriteErrorMsg(readListener, path, "模板", DemoData.class);

//        ExcelUtils.addErrorInfo(path, path, "模板", readListener.getErrorInfoList());
        System.out.println("次耗时：" + (System.currentTimeMillis() - startTime) / 1000 + "s");
        System.out.println("end");
    }


    @SneakyThrows
    public static void write(OutputStream outputStream, List<CellErrorInfo> cellErrorInfoList) {
        List classDataList = getDataList();
        EasyExcel.write(outputStream)
                .sheet("模板")
//                .head(getHead())
                .head(DemoData.class)
                .registerWriteHandler(new ExtendColumnHandler<DemoData>(classDataList))
                .registerWriteHandler(new HeadStyleHandler())
//                .registerWriteHandler(new ErrorInfoCommentHandler(cellErrorInfoList))
//                .registerWriteHandler(new DropDownBoxSheetHandler(getDropDownBoxInfo()))
//                .excludeColumnFiledNames(Arrays.asList("dest"))
//                .doWrite(Collections.emptyList());
//                .doWrite(getDataList());
                .doWrite(classDataList);
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

    public static List<List<String>> getHead() {
        List<List<String>> list = new ArrayList<List<String>>();
        List<String> head0 = new ArrayList<String>();
        head0.add("字符串");
        List<String> head1 = new ArrayList<String>();
        head1.add("数字");
        List<String> head2 = new ArrayList<String>();
        head2.add("dest");
        list.add(head0);
        list.add(head1);
        list.add(head2);
        return list;
    }

    public static List<ValidationInfo> getDropDownBoxInfo() {
        List<ValidationInfo> validationInfoList = new ArrayList<>();
        validationInfoList.add(ValidationInfo.buildColumn("SRC", "是", "否"));
        validationInfoList.add(ValidationInfo.buildColumn(1,"可以", "不可以"));
        validationInfoList.add(ValidationInfo.bulidRow(2, "中", "不中"));
        validationInfoList.add(ValidationInfo.buildPrecise(3, "媒体CODE", "不中"));
        return validationInfoList;
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
            demoData.setId(Long.valueOf(i));
            demoData.setDest("dest" + i);
            demoData.setSrc("src" + i);
            demoData.setFunc("func" + i);

            Map<String, String> map = new LinkedHashMap<>();
            map.put("自定义1", "12");
            map.put("自定义2", "jfak");
            map.put("自定义3", "集分宝");
            demoData.setDynamicColumnMap(map);
            dataList.add(demoData);
        }
        return dataList;
    }

    private static List<List<Object>> getDataList1(List<List<String>> heads) {
        List<List<Object>> dataList = new ArrayList<>();
        for (int rowIndex = 0; rowIndex < 10; rowIndex++) {
            for (int j = 0; j < heads.size(); j++) {
                List<Object> data;
                if (rowIndex >= dataList.size()) {
                    data = new ArrayList<>();
                    dataList.add(rowIndex, data);
                } else {
                    data = dataList.get(rowIndex);
                }
                data.add(heads.get(j).get(0) + rowIndex);
            }
        }
        return dataList;
    }

}

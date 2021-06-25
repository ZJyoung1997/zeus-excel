package com.jz.zeus.excel.test;

import cn.hutool.core.collection.ListUtil;
import cn.hutool.core.date.DateUtil;
import cn.hutool.core.date.TimeInterval;
import cn.hutool.core.lang.Console;
import com.alibaba.excel.read.listener.ReadListener;
import com.alibaba.excel.util.IoUtils;
import com.jz.zeus.excel.CellErrorInfo;
import com.jz.zeus.excel.ZeusExcel;
import com.jz.zeus.excel.read.listener.ExcelReadListener;
import com.jz.zeus.excel.test.data.DemoData;
import com.jz.zeus.excel.test.listener.DemoReadListener;
import lombok.SneakyThrows;

import java.io.ByteArrayInputStream;
import java.io.FileInputStream;
import java.util.*;
import java.util.stream.Collectors;
import java.util.stream.Stream;

/**
 * @Author JZ
 * @Date 2021/4/2 11:36
 */
public class ExcelReadTest {

//    private static String path = "C:\\Users\\Administrator\\Desktop\\254.xlsx";
    private static String path = "C:\\Users\\User\\Desktop\\254.xlsx";

//    private static String path = "C:\\Users\\User\\Desktop\\2545.xlsx";
//    private static String path = "C:\\Users\\Administrator\\Desktop\\2545.xlsx";

    @SneakyThrows
    public static void main(String[] args) {
        Console.log("解析Excel前内存：{}M", (Runtime.getRuntime().totalMemory() - Runtime.getRuntime().freeMemory())/(1024*1024));
        TimeInterval timer = DateUtil.timer();

        ExcelReadListener<?> readListener = new DemoReadListener();

//        byte[] bytes = IoUtils.toByteArray(new FileInputStream(path));
//        System.out.println("解析为字节后内存："+(Runtime.getRuntime().totalMemory() - Runtime.getRuntime().freeMemory())/(1024*1024)+"M");

        read(path, readListener);

        List<CellErrorInfo> errorInfos = readListener.getErrorInfoMap().values()
                .stream().flatMap(e -> e.stream()).collect(Collectors.toList());
        byte[] bytes = IoUtils.toByteArray(new FileInputStream(path));
        ZeusExcel.write(path)
                .withTemplate(new ByteArrayInputStream(bytes))
                .sheet(0)
                .errorInfos(errorInfos)
                .doWrite(DemoData.class, Collections.emptyList());


        Console.log("读取耗时：{}s", timer.intervalSecond());
        Console.log("解析Excel后内存：{}M", (Runtime.getRuntime().totalMemory() - Runtime.getRuntime().freeMemory())/(1024*1024));
    }

    public static void read(String path, ReadListener readListener) {
        ZeusExcel.read(path)
                .sheet("模板")
                .head(DemoData.class)
                .registerReadListener(readListener)
                .doRead();
    }

    public static List<List<String>> getHead() {
        List<List<String>> list = new ArrayList<List<String>>();
        List<String> head0 = new ArrayList<String>();
        head0.add("ID");
        List<String> head1 = new ArrayList<String>();
        head1.add("SRC");
        List<String> head2 = new ArrayList<String>();
        head2.add("DEST");
        List<String> head3 = new ArrayList<String>();
        head3.add("FUNC");
        list.add(head0);
        list.add(head1);
        list.add(head2);
//        list.add(head3);
        return list;
    }

}

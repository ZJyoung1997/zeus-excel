package com.jz.zeus.excel.test;

import cn.hutool.core.collection.ListUtil;
import cn.hutool.core.lang.Console;
import cn.hutool.core.util.HexUtil;
import lombok.SneakyThrows;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.DefaultIndexedColorMap;
import org.apache.poi.xssf.usermodel.XSSFColor;

import java.awt.*;
import java.util.List;
import java.util.stream.Collectors;

/**
 * @Author JZ
 * @Date 2021/4/12 16:31
 */
public class Test {

    @SneakyThrows
    public static void main(String[] args) {
        for (IndexedColors value : IndexedColors.values()) {
            Console.log(value.index);
        }
        Console.log(HexUtil.encodeHexStr(new DefaultIndexedColorMap().getRGB(1)));
    }


}

package com.jz.zeus.excel.test;

import cn.hutool.core.collection.ListUtil;
import cn.hutool.core.lang.Console;
import lombok.SneakyThrows;

import java.util.List;
import java.util.stream.Collectors;

/**
 * @Author JZ
 * @Date 2021/4/12 16:31
 */
public class Test {

    @SneakyThrows
    public static void main(String[] args) {
        List<String> list = ListUtil.toList("fjkd", null, "8388", null, "福晶科技");
        list = list.stream().collect(Collectors.toList());
        Console.log(list);
    }


}

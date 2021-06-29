package com.jz.zeus.excel.test;

import cn.hutool.core.collection.ListUtil;
import com.jz.zeus.excel.CellErrorInfo;
import lombok.SneakyThrows;

import java.util.*;
import java.util.stream.Collectors;

/**
 * @Author JZ
 * @Date 2021/4/12 16:31
 */
public class Test {

    public static void main(String[] args) {
        Map<Integer, List<CellErrorInfo>> map = new HashMap<>();
        map.computeIfAbsent(1, ArrayList::new)
                .addAll(ListUtil.toList(CellErrorInfo.buildByField(1, "jfk", "fjkj")));
        map.computeIfAbsent(2, ArrayList::new)
                .addAll(ListUtil.toList(CellErrorInfo.buildByField(1, "jfk", "fjkj")));
        List<CellErrorInfo> list = map.values().stream().flatMap(Collection::stream)
                .collect(Collectors.toList());
    }

}

package com.jz.zeus.excel.test;

import cn.hutool.core.collection.CollUtil;
import cn.hutool.core.util.ArrayUtil;
import com.jz.zeus.excel.CellErrorInfo;
import com.jz.zeus.excel.ValidationInfo;
import lombok.SneakyThrows;

import java.util.ArrayList;
import java.util.Collections;
import java.util.List;

/**
 * @Author JZ
 * @Date 2021/4/12 16:31
 */
public class Test {

    @SneakyThrows
    public static void main(String[] args) {
        CellErrorInfo c1 = CellErrorInfo.buildByColumnIndex(0, 2, "急急急", "jiijie");
        CellErrorInfo c2 = c1.clone();
        c2.setErrorMsgs(CollUtil.newArrayList("jfkdj"));
        System.out.println(c2);
    }

}

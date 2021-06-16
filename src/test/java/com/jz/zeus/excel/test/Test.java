package com.jz.zeus.excel.test;

import cn.hutool.core.lang.Console;
import cn.hutool.core.util.HexUtil;
import com.jz.zeus.excel.ValidationInfo;
import com.jz.zeus.excel.test.data.A;
import lombok.SneakyThrows;
import org.apache.poi.xssf.usermodel.DefaultIndexedColorMap;

/**
 * @Author JZ
 * @Date 2021/4/12 16:31
 */
public class Test {

    @SneakyThrows
    public static void main(String[] args) {
        ValidationInfo info = ValidationInfo.buildColumnByField(A::getName, "jkjk");
        Console.log(HexUtil.encodeHexStr(new DefaultIndexedColorMap().getRGB(1)));
    }

}

package com.jz.zeus.excel.test;

import com.jz.zeus.excel.util.ValidatorUtils;

/**
 * @author:JZ
 * @date:2021/3/24
 */
public class ValidatorTest {

    public static void main(String[] args) {
        DemoData demoData = new DemoData();
        ValidatorUtils.validate(demoData);
    }

}

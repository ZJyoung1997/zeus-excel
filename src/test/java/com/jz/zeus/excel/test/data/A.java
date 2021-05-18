package com.jz.zeus.excel.test.data;

import lombok.Data;

import java.util.List;

/**
 * @Author JZ
 * @Date 2021/5/18 12:13
 */
@Data
public class A {

    private Long id;

    private String name;

    private List<B> bList;

    private B b;

}

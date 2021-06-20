package com.jz.zeus.excel.annotation;

import com.alibaba.excel.annotation.ExcelIgnore;
import com.jz.zeus.excel.read.listener.ExcelReadListener;

import java.lang.annotation.*;

/**
 * 标注该注解的方法，表示其为动态字段，其类型应该为Map<String, String>，key 为表头、value 为单元格数据
 * 在 {@link ExcelReadListener} 中会将Excel读取到的动态数据存入其中。
 * 该注解应配合 {@link ExcelIgnore} 注解使用，以使其不被作为表头写入Excel。
 *
 * @Author JZ
 * @Date 2021/4/12 14:03
 */
@Inherited
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
public @interface ExtendColumn {}

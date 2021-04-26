package com.jz.zeus.excel;

import com.alibaba.excel.metadata.property.ColumnWidthProperty;
import com.alibaba.excel.metadata.property.FontProperty;
import com.alibaba.excel.metadata.property.LoopMergeProperty;
import com.alibaba.excel.metadata.property.StyleProperty;
import com.jz.zeus.excel.annotation.ValidationData;
import lombok.Data;

import java.lang.reflect.Field;

/**
 * @Author JZ
 * @Date 2021/4/8 17:52
 */
@Data
public class FieldInfo {

    private Field field;

    /**
     * 字段名
     */
    private String fieldName;

    private String headName;

    private Integer headColumnIndex;

    /**
     * 下拉框信息
     */
    private ValidationData validationData;

    /**
     * 是否为动态列
     */
    private boolean isExtendColumn;

    private FontProperty headFontProperty;

    private StyleProperty headStyleProperty;

    private FontProperty contentFontProperty;

    private StyleProperty contentStyleProperty;

    private ColumnWidthProperty columnWidthProperty;

    private LoopMergeProperty loopMergeProperty;

}

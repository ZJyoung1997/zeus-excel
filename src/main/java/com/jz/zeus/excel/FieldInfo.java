package com.jz.zeus.excel;

import lombok.Data;

/**
 * @Author JZ
 * @Date 2021/4/8 17:52
 */
@Data
public class FieldInfo {

    /**
     * 字段名
     */
    private String fieldName;

    private String headName;

    private Integer headColumnIndex;

    /**
     * 下拉框内容
     */
    private String[] dropDownBoxOptions;

    /**
     * 下拉框需要填充的行数
     */
    private Integer dropDownBoxRowNum;

}

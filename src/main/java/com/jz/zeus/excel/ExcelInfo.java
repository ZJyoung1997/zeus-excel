package com.jz.zeus.excel;

import lombok.Data;

import java.util.List;

/**
 * @author:JZ
 * @date:2021/3/18
 */
@Data
public class ExcelInfo {

    /**
     * excel 名称
     */
    private String excelName;

    /**
     * 标签页名称列表
     */
    private List<String> sheetNames;

    /**
     * 需要写入的列（其表头和对应数据都会写入Excel），值为对应表头名
     */
    private List<String> includeColumnNames;

    /**
     * 不需要写入的列（其表头和对应数据不会写入Excel），值为对应表头名
     */
    private List<String> excludeColumnNames;

}

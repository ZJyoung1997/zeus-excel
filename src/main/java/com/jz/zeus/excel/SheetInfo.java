package com.jz.zeus.excel;

import cn.hutool.core.collection.CollectionUtil;
import lombok.Data;

import java.util.Collection;
import java.util.HashSet;
import java.util.List;

/**
 * @author:JZ
 * @date:2021/3/21
 */
@Data
public class SheetInfo<T> {

    /**
     * 标签页名称列表
     */
    private String name;

    /**
     * 标签页索引
     */
    private Integer index;

    /**
     * 需要写入的列（其表头和对应数据都会写入Excel），值为对应表头名
     */
    private Collection<String> includeHeaderNames;

    /**
     * 不需要写入的列（其表头和对应数据不会写入Excel），值为对应表头名
     */
    private Collection<String> excludeHeaderNames;

    private List<T> dataList;

    public void addIncludeHeaderName(String headerName) {
        if (CollectionUtil.isEmpty(includeHeaderNames)) {
            includeHeaderNames = new HashSet<>();
        }
        includeHeaderNames.add(headerName);
    }

    public void addIncludeHeaderName(Collection<String> headerNames) {
        if (CollectionUtil.isEmpty(includeHeaderNames)) {
            includeHeaderNames = new HashSet<>();
        }
        includeHeaderNames.addAll(headerNames);
    }

    public void addExcludeHeaderName(String headerName) {
        if (CollectionUtil.isEmpty(excludeHeaderNames)) {
            excludeHeaderNames = new HashSet<>();
        }
        excludeHeaderNames.add(headerName);
    }

    public void addExcludeHeaderName(Collection<String> headerNames) {
        if (CollectionUtil.isEmpty(excludeHeaderNames)) {
            excludeHeaderNames = new HashSet<>();
        }
        excludeHeaderNames.addAll(headerNames);
    }

}

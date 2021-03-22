package com.jz.zeus.excel.read.listener;

import cn.hutool.core.collection.CollUtil;
import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.metadata.CellData;
import com.alibaba.excel.metadata.CellExtra;
import com.alibaba.excel.read.listener.ReadListener;
import com.jz.zeus.excel.CellErrorInfo;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

import java.util.*;
import java.util.stream.Collectors;

/**
 * @Author JZ
 * @Date 2021/3/22 14:31
 */
@NoArgsConstructor
public abstract class AbstractExcelReadListener<T> implements ReadListener<T> {

    /**
     * 表头是否有误
     * false 正常、true 表头错误
     */
    @Getter
    protected boolean headError = false;

    /**
     * 为true时，所有数据加载到 dataList 且执行完 verify() 方法后对数据进行保存，
     * 这意味着 batchSaveNum 将无效
     */
    @Setter
    protected boolean lastSave = false;

    /**
     * 批量保存数量
     */
    @Setter
    protected int batchSaveNum = 500;

    /**
     * 错误信息，可以是表头或数据的错误信息
     * key 为行索引，value 为单元格错误信息
     */
    public Map<Integer, List<CellErrorInfo>> rowErrorInfoMap;

    /**
     * Excel 中读取到的数据
     */
    protected List<T> dataList = new ArrayList<>();

    public AbstractExcelReadListener(boolean lastSave) {
        this.lastSave = lastSave;
    }

    public AbstractExcelReadListener(int batchSaveNum) {
        this.batchSaveNum = batchSaveNum;
    }

    public AbstractExcelReadListener(boolean lastSave, int batchSaveNum) {
        this.lastSave = lastSave;
        this.batchSaveNum = batchSaveNum;
    }

    /**
     * 保存 dataList 中的数据
     */
    protected abstract void save(AnalysisContext analysisContext);

    /**
     * 校验 dataList 中的数据
     */
    protected List<CellErrorInfo> verify(AnalysisContext analysisContext) {
        return Collections.emptyList();
    };

    /**
     * 校验表头是否正常
     * 可调用 ConverterUtils.convertToStringMap(headMap, context) 方法将表头转化为对应map
     * @return 正常 true、异常 false
     */
    public boolean headCheck(Map<Integer, CellData> headMap, AnalysisContext context) {
        return true;
    }

    @Override
    public void invoke(T data, AnalysisContext analysisContext) {
        dataList.add(data);
        if (!lastSave && dataList.size() >= batchSaveNum) {
            addRowErrorInfo(verify(analysisContext));
            save(analysisContext);
            dataList.clear();
        }
    }

    @Override
    public void doAfterAllAnalysed(AnalysisContext analysisContext) {
        if (dataList.size() > 0) {
            addRowErrorInfo(verify(analysisContext));
            save(analysisContext);
            dataList.clear();
        }
    }

    @Override
    public void invokeHead(Map<Integer, CellData> map, AnalysisContext analysisContext) {
        if (!headCheck(map, analysisContext)) {
            headError = false;
        }
    }

    /**
     * 当表头有误时，停止Excel的读取
     * @param context
     * @return
     */
    @Override
    public boolean hasNext(AnalysisContext context) {
        if (headError) {
            return false;
        }
        return true;
    }

    protected void addRowErrorInfo(List<CellErrorInfo> cellErrorInfoList) {
        if (CollUtil.isEmpty(cellErrorInfoList)) {
            return;
        }
        if (CollUtil.isEmpty(rowErrorInfoMap)) {
            rowErrorInfoMap = cellErrorInfoList.stream()
                    .collect(Collectors.groupingBy(CellErrorInfo::getRowIndex));
            return;
        }
        cellErrorInfoList.stream().collect(Collectors.groupingBy(CellErrorInfo::getRowIndex))
                .forEach((rowIndex, value) -> {
                    List<CellErrorInfo> cellErrorInfos = rowErrorInfoMap.get(rowIndex);
                    if (CollUtil.isEmpty(cellErrorInfos)) {
                        cellErrorInfos = new ArrayList<>();
                        rowErrorInfoMap.put(rowIndex, cellErrorInfos);
                    }
                    cellErrorInfos.addAll(value);
                });
    }

    @Override
    public void onException(Exception e, AnalysisContext analysisContext) {}

    @Override
    public void extra(CellExtra cellExtra, AnalysisContext analysisContext) {}

}

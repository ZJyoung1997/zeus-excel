package com.jz.zeus.excel.read.listener;

import cn.hutool.core.collection.CollUtil;
import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.metadata.CellData;
import com.alibaba.excel.metadata.CellExtra;
import com.alibaba.excel.metadata.Head;
import com.alibaba.excel.read.listener.ReadListener;
import com.alibaba.excel.read.metadata.holder.ReadRowHolder;
import com.alibaba.excel.read.metadata.property.ExcelReadHeadProperty;
import com.jz.zeus.excel.CellErrorInfo;
import com.jz.zeus.excel.util.ValidatorUtils;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

import java.util.*;

/**
 * @Author JZ
 * @Date 2021/3/22 14:31
 */
@NoArgsConstructor
public abstract class AbstractExcelReadListener<T> implements ReadListener<T> {

    /**
     * 开启hibernate注解校验
     * true 开启、false 关闭
     */
    protected boolean enabledAnnotationValidation = true;

    /**
     * 表头是否有误
     * false 正常、true 表头错误
     */
    @Getter
    protected boolean headError = false;

    /**
     * 为true时，所有数据加载到 dataList 且执行完 verify() 方法后对数据进行处理，
     * 这意味着 batchSaveNum 将无效
     */
    @Setter
    protected boolean lastHandleData = false;

    /**
     * 批量保存数量
     */
    @Setter
    protected int batchSaveNum = 500;

    /**
     * 错误信息，可以是表头或数据的错误信息
     * key 为行索引，value 为单元格错误信息
     */
    @Getter
    protected List<CellErrorInfo> rowErrorInfoList;

    /**
     * T 中字段名与列索引映射
     * key T 中字段名、value 列索引
     */
    protected Map<String, Integer> fieldColumnIndexMap;

    /**
     * Excel 中读取到的数据
     */
    protected List<T> dataList = new ArrayList<>();

    public AbstractExcelReadListener(boolean lastHandleData) {
        this(lastHandleData, null, null);
    }

    public AbstractExcelReadListener(int batchSaveNum) {
        this(null, null, batchSaveNum);
    }

    public AbstractExcelReadListener(Boolean lastHandleData, Boolean enabledAnnotationValidation, Integer batchSaveNum) {
        if (lastHandleData != null) {
            this.lastHandleData = lastHandleData;
        }
        if (enabledAnnotationValidation != null) {
            this.enabledAnnotationValidation = enabledAnnotationValidation;
        }
        if (batchSaveNum != null) {
            this.batchSaveNum = batchSaveNum;
        }
    }

    /**
     * 保存 dataList 中的数据
     */
    protected abstract void dataHandle(AnalysisContext analysisContext, Integer currentRowIndex);

    /**
     * 校验 dataList 中的数据
     */
    protected List<CellErrorInfo> verify(AnalysisContext analysisContext, Integer currentRowIndex) {
        return Collections.emptyList();
    };

    /**
     * 校验表头是否正常
     * 可调用 ConverterUtils.convertToStringMap(headMap, context) 方法将表头转化为对应map
     * @return 正常 true、异常 false
     */
    protected List<CellErrorInfo> headCheck(Map<Integer, CellData> headMap, AnalysisContext context) {
        return Collections.emptyList();
    }

    @Override
    public void invoke(T data, AnalysisContext analysisContext) {
        annotationValidation(data, analysisContext.readRowHolder());
        dataList.add(data);
        if (!lastHandleData && dataList.size() >= batchSaveNum) {
            Integer currentRowIndex = analysisContext.readRowHolder().getRowIndex();
            addRowErrorInfo(verify(analysisContext, currentRowIndex));
            dataHandle(analysisContext, currentRowIndex);
            dataList.clear();
        }
    }

    @Override
    public void doAfterAllAnalysed(AnalysisContext analysisContext) {
        if (dataList.size() > 0) {
            Integer currentRowIndex = analysisContext.readRowHolder().getRowIndex();
            addRowErrorInfo(verify(analysisContext, currentRowIndex));
            dataHandle(analysisContext, currentRowIndex);
            dataList.clear();
        }
    }

    @Override
    public void invokeHead(Map<Integer, CellData> headMap, AnalysisContext analysisContext) {
        List<CellErrorInfo> cellErrorInfoList = headCheck(headMap, analysisContext);
        if (CollUtil.isNotEmpty(cellErrorInfoList)) {
            addRowErrorInfo(cellErrorInfoList);
            headError = false;
        }
        ExcelReadHeadProperty excelHeadPropertyData = analysisContext.readSheetHolder().excelReadHeadProperty();
        Map<Integer, Head> headMapData = excelHeadPropertyData.getHeadMap();
        if (CollUtil.isEmpty(headMapData)) {
            fieldColumnIndexMap = Collections.emptyMap();
            return;
        }
        fieldColumnIndexMap = new HashMap<>();
        headMapData.values().forEach(head -> {
            fieldColumnIndexMap.put(head.getFieldName(), head.getColumnIndex());
        });
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

    @Override
    public void onException(Exception e, AnalysisContext analysisContext) {}

    @Override
    public void extra(CellExtra cellExtra, AnalysisContext analysisContext) {}

    protected void addRowErrorInfo(List<CellErrorInfo> cellErrorInfos) {
        if (CollUtil.isEmpty(cellErrorInfos)) {
            return;
        }
        if (CollUtil.isEmpty(rowErrorInfoList)) {
            rowErrorInfoList = new ArrayList<>();
        }
        rowErrorInfoList.addAll(cellErrorInfos);
    }

    protected void annotationValidation(T data, ReadRowHolder readRowHolder) {
        if (!enabledAnnotationValidation) {
            return;
        }
        Map<String, List<String>> errorMessageMap = ValidatorUtils.validate(data);
        if (CollUtil.isEmpty(errorMessageMap)) {
            return;
        }
        List<CellErrorInfo> errorMesInfoList = new ArrayList<>();
        errorMessageMap.forEach((fieldName, errorMessages) -> {
            errorMesInfoList.add(new CellErrorInfo(readRowHolder.getRowIndex(), fieldColumnIndexMap.get(fieldName))
                                .addErrorMsg(errorMessages));

        });
        addRowErrorInfo(errorMesInfoList);
    }

    public boolean hasError() {
        return CollUtil.isNotEmpty(rowErrorInfoList);
    }

}

package com.jz.zeus.excel.read.listener;

import cn.hutool.core.collection.CollUtil;
import cn.hutool.core.util.StrUtil;
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
    private boolean enabledAnnotationValidation = true;

    /**
     * 表头是否有误
     * false 正常、true 表头错误
     */
    @Getter
    private boolean headError;

    /**
     * 为true时，所有数据加载到 dataList 且执行完 verify() 方法后对数据进行处理，
     * 这意味着 batchSaveNum 将无效
     */
    @Setter
    private boolean lastHandleData = false;

    /**
     * 批量保存数量
     */
    @Setter
    private int batchSaveNum = 500;

    /**
     * 表头错误信息
     */
    @Setter
    @Getter
    private String headErrorMsg;

    /**
     * 数据的错误信息（不包含表头错误信息）
     * key 为行索引，value 为单元格错误信息
     */
    @Getter
    protected List<CellErrorInfo> errorInfoList;

    /**
     * T 中字段名与列索引映射
     * key T 中字段名、value 列索引
     */
    private Map<String, Integer> fieldColumnIndexMap;

    /**
     * 表头与列索引映射
     * key 表头、value 表头对应列索引
     */
    private Map<String, Integer> headNameIndexMap;

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
     * 所有数据处理完毕之后触发的操作
     */
    protected abstract void doAfterAllDataHandle(AnalysisContext analysisContext);

    /**
     * 保存 dataList 中的数据
     */
    protected abstract void dataHandle(AnalysisContext analysisContext, Integer currentRowIndex);

    /**
     * 校验 dataList 中的数据
     */
    protected abstract void verify(AnalysisContext analysisContext, Integer currentRowIndex);

    /**
     * 校验表头是否正常
     * 可调用 ConverterUtils.convertToStringMap(headMap, context) 方法将表头转化为对应map
     * @return 正常 true、异常 false
     */
    protected abstract void headCheck(Map<Integer, CellData> headMap, AnalysisContext context);

    @Override
    public void invoke(T data, AnalysisContext analysisContext) {
        annotationValidation(data, analysisContext.readRowHolder());
        dataList.add(data);
        if (!lastHandleData && dataList.size() >= batchSaveNum) {
            Integer currentRowIndex = analysisContext.readRowHolder().getRowIndex();
            verify(analysisContext, currentRowIndex);
            dataHandle(analysisContext, currentRowIndex);
            dataList.clear();
        }
    }

    @Override
    public void doAfterAllAnalysed(AnalysisContext analysisContext) {
        if (dataList.size() > 0) {
            Integer currentRowIndex = analysisContext.readRowHolder().getRowIndex();
            verify(analysisContext, currentRowIndex);
            dataHandle(analysisContext, currentRowIndex);
            dataList.clear();
        }
        doAfterAllDataHandle(analysisContext);
    }

    @Override
    public void invokeHead(Map<Integer, CellData> headMap, AnalysisContext analysisContext) {
        headError = false;
        headCheck(headMap, analysisContext);
        if (StrUtil.isNotBlank(this.headErrorMsg)) {
            headError = true;
        }
        ExcelReadHeadProperty excelHeadPropertyData = analysisContext.readSheetHolder().excelReadHeadProperty();
        Map<Integer, Head> headMapData = excelHeadPropertyData.getHeadMap();
        if (CollUtil.isEmpty(headMapData)) {
            fieldColumnIndexMap = Collections.emptyMap();
            return;
        }
        fieldColumnIndexMap = new HashMap<>();
        headNameIndexMap = new HashMap<>();
        headMapData.values().forEach(head -> {
            Integer columnIndex = head.getColumnIndex();
            fieldColumnIndexMap.put(head.getFieldName(), columnIndex);
            head.getHeadNameList().forEach(headName -> {
                headNameIndexMap.put(headName, columnIndex);
            });
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

    protected void addErrorInfo(Integer rowIndex, String headName, String... errorMessages) {
        Integer columnIndex = headNameIndexMap.get(headName);
        if (columnIndex == null) {
            return;
        }
        errorInfoList.add(new CellErrorInfo(rowIndex, columnIndex, errorMessages));
    }

    protected void addErrorInfo(List<CellErrorInfo> cellErrorInfos) {
        if (CollUtil.isEmpty(cellErrorInfos)) {
            return;
        }
        if (CollUtil.isEmpty(errorInfoList)) {
            this.errorInfoList = new ArrayList<>();
        }
        this.errorInfoList.addAll(cellErrorInfos);
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
        addErrorInfo(errorMesInfoList);
    }

    /**
     * 判断是否存在表头错误
     * @return true 存在表头错误、false 不存在表头错误
     */
    public boolean hasHeadError() {
        return this.headError;
    }

    /**
     * 判断是否存在数据错误
     * @return true 存在数据错误、false 不存在数据错误
     */
    public boolean hasDataError() {
        return CollUtil.isNotEmpty(errorInfoList);
    }

}

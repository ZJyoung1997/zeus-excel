package com.jz.zeus.excel.read.listener;

import cn.hutool.core.collection.CollUtil;
import cn.hutool.core.text.CharSequenceUtil;
import cn.hutool.core.util.ReflectUtil;
import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.enums.HeadKindEnum;
import com.alibaba.excel.metadata.Cell;
import com.alibaba.excel.metadata.CellData;
import com.alibaba.excel.metadata.CellExtra;
import com.alibaba.excel.metadata.Head;
import com.alibaba.excel.read.listener.ReadListener;
import com.alibaba.excel.read.metadata.holder.ReadRowHolder;
import com.alibaba.excel.read.metadata.property.ExcelReadHeadProperty;
import com.alibaba.excel.util.ConverterUtils;
import com.jz.zeus.excel.CellErrorInfo;
import com.jz.zeus.excel.FieldInfo;
import com.jz.zeus.excel.exception.DataConvertException;
import com.jz.zeus.excel.interfaces.FieldGetter;
import com.jz.zeus.excel.util.ClassUtils;
import lombok.AccessLevel;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

import java.lang.reflect.Field;
import java.util.*;
import java.util.stream.Collectors;

@NoArgsConstructor
public abstract class ExcelReadListener<T> implements ReadListener<T> {

    /**
     * 批量处理数据数量
     */
    private int batchHandleNum = 200;

    /**
     * 终止标志，true 终止Excel读取流程、false Excel读取流程正常执行
     */
    @Setter(AccessLevel.PROTECTED)
    private boolean terminated;

    /**
     * 扩展列起始索引，若没有扩展列则为null
     */
    private Integer extendColumnBeginIndex;

    /**
     * 扩展列结束索引，若没有扩展列则为null
     */
    private Integer extendColumnEndIndex;

    /**
     * 扩展列的字段信息
     */
    private Field extendField;

    /**
     * 当前sheet中读取到的表头
     */
    private Map<Integer, String> currentSheetHeads;

    /**
     * 当前sheet表头配置信息
     */
    private Map<Integer, Head> currentSheetHeadConfig;

    /**
     * Excel 中读取到的数据
     * key 行索引，value 读取到的数据
     */
    private Map<Integer, T> dataMap;

    /**
     * 数据的错误信息
     * key 为行索引，value 为该行中单元格的错误信息
     */
    @Getter
    private Map<Integer, List<CellErrorInfo>> errorInfoMap;

    public ExcelReadListener(int batchHandleNum) {
        this.batchHandleNum = batchHandleNum;
    }

    /**
     * 对sheet表头进行处理
     * @param headMap
     * @param analysisContext
     */
    protected abstract void headHandle(Map<Integer, String> headMap, AnalysisContext analysisContext);

    /**
     * 对读取到的数据进行处理
     * @param dataMap                   key 为行索引，value 为 Excel中该行数据
     * @param analysisContext
     */
    protected abstract void dataHandle(Map<Integer, T> dataMap, AnalysisContext analysisContext);

    /**
     * 所有数据处理完毕之后触发的操作，如果想要所有数据校验完毕后在对数据进行操作
     * 可以对 {@link #dataHandle} 进行一个空的实现
     */
    protected abstract void doAfterAllDataHandle(AnalysisContext analysisContext);

    @Override
    public void invokeHead(Map<Integer, CellData> headMap, AnalysisContext context) {
        init(headMap, context);
        headHandle(currentSheetHeads, context);
    }

    @Override
    public void invoke(T data, AnalysisContext context) {
        ReadRowHolder readRowHolder = context.readRowHolder();
        Integer rowIndex = readRowHolder.getRowIndex();
        setExtendColumnData(data, readRowHolder);
        dataMap.put(rowIndex, data);
        if (dataMap.size() >= batchHandleNum) {
            dataHandle(dataMap, context);
            dataMap.clear();
        }
    }

    @Override
    public void doAfterAllAnalysed(AnalysisContext context) {
        if (dataMap.size() > 0) {
            dataHandle(dataMap, context);
            dataMap.clear();
        }
        doAfterAllDataHandle(context);
    }

    @Override
    public boolean hasNext(AnalysisContext context) {
        return !terminated;
    }

    @Override
    public void onException(Exception exception, AnalysisContext context) {
        if (exception instanceof DataConvertException) {
            List<CellErrorInfo> errorInfos = ((DataConvertException) exception).getCellErrorInfos();
            if (CollUtil.isNotEmpty(errorInfos)) {
                errorInfoMap.putAll(errorInfos.stream().collect(Collectors.groupingBy(CellErrorInfo::getRowIndex)));
            }
        } else {
            throw new RuntimeException(exception);
        }
    }

    protected void addErrorInfo(int rowIndex, String headName, String... errorMsgs) {
        List<CellErrorInfo> rowErrorInfoList = errorInfoMap.computeIfAbsent(rowIndex, k -> new ArrayList<>());
        rowErrorInfoList.add(CellErrorInfo.buildByHead(rowIndex, headName, errorMsgs));
    }

    protected <T, R> void addErrorInfo(int rowIndex, FieldGetter<T, R> fieldGetter, String... errorMsgs) {
        List<CellErrorInfo> rowErrorInfoList = errorInfoMap.computeIfAbsent(rowIndex, k -> new ArrayList<>());
        rowErrorInfoList.add(CellErrorInfo.buildByField(rowIndex, fieldGetter, errorMsgs));
    }

    protected void addErrorInfo(int rowIndex, int columnIndex, String... errorMsgs) {
        List<CellErrorInfo> rowErrorInfoList = errorInfoMap.computeIfAbsent(rowIndex, k -> new ArrayList<>());
        Head head = currentSheetHeadConfig.get(columnIndex);
        String fieldName = head.getFieldName();
        if (CharSequenceUtil.isNotBlank(fieldName)) {
            rowErrorInfoList.add(CellErrorInfo.buildByField(rowIndex, fieldName, errorMsgs));
        } else {
            rowErrorInfoList.add(CellErrorInfo.buildByColumnIndex(rowIndex, columnIndex, errorMsgs));
        }
    }

    private void init(Map<Integer, CellData> headMap, AnalysisContext context) {
        terminated = false;
        extendField = null;
        extendColumnBeginIndex = null;
        extendColumnEndIndex = null;
        errorInfoMap = new HashMap<>();
        dataMap = new HashMap<>(batchHandleNum);
        currentSheetHeads = ConverterUtils.convertToStringMap(headMap, context);
        int maxColumnIndex = currentSheetHeads.keySet().stream()
                .max(Integer::compareTo).orElse(-1);

        ExcelReadHeadProperty excelHeadPropertyData = context.readSheetHolder().excelReadHeadProperty();
        HeadKindEnum headKind = excelHeadPropertyData.getHeadKind();
        currentSheetHeadConfig = excelHeadPropertyData.getHeadMap();
        if (HeadKindEnum.CLASS == headKind) {
            ClassUtils.getClassFieldInfo(excelHeadPropertyData.getHeadClazz())
                    .stream().filter(FieldInfo::isExtendColumn).findFirst()
                    .ifPresent(fieldInfo -> {
                        extendField = fieldInfo.getField();
                        extendColumnBeginIndex = currentSheetHeadConfig.keySet()
                                .stream().max(Integer::compareTo).orElse(null);
                        if (extendColumnBeginIndex != null) {
                            extendColumnBeginIndex++;
                        }
                        extendColumnEndIndex = maxColumnIndex;
                    });
        }
    }

    /**
     * 将Excel中扩展列的数据存入对象中
     * @param data
     * @param readRowHolder
     */
    private void setExtendColumnData(T data, ReadRowHolder readRowHolder) {
        if (extendColumnBeginIndex == null || extendColumnEndIndex == null) {
            return;
        }
        Map<Integer, Cell> cellDataMap = readRowHolder.getCellMap();
        if (CollUtil.isEmpty(cellDataMap)) {
            return;
        }
        Map<String, String> extendData = new LinkedHashMap<>();
        for (int i = extendColumnBeginIndex; i <= extendColumnEndIndex; i++) {
            Cell cell = cellDataMap.get(i);
            String extendHeadName = currentSheetHeads.get(i);
            if (cell != null) {
                extendData.put(extendHeadName, cell.toString());
            } else {
                extendData.put(extendHeadName, null);
            }
        }

        if (Map.class == extendField.getType()) {
            ReflectUtil.setFieldValue(data, extendField, extendData);
        }
    }

    @Override
    public void extra(CellExtra extra, AnalysisContext context) {}

}

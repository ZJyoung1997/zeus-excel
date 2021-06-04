package com.jz.zeus.excel.util;

import cn.hutool.core.collection.CollUtil;
import cn.hutool.core.collection.ListUtil;
import cn.hutool.core.util.StrUtil;
import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.annotation.write.style.*;
import com.alibaba.excel.metadata.property.ColumnWidthProperty;
import com.alibaba.excel.metadata.property.FontProperty;
import com.alibaba.excel.metadata.property.LoopMergeProperty;
import com.alibaba.excel.metadata.property.StyleProperty;
import com.jz.zeus.excel.FieldInfo;
import com.jz.zeus.excel.ValidationInfo;
import com.jz.zeus.excel.annotation.ExtendColumn;
import com.jz.zeus.excel.annotation.ValidationData;
import lombok.experimental.UtilityClass;
import org.hibernate.validator.internal.util.ConcurrentReferenceHashMap;

import java.lang.reflect.Field;
import java.util.*;
import java.util.stream.Collectors;

/**
 * @Author JZ
 * @Date 2021/4/8 17:20
 */
@UtilityClass
public class ClassUtils {

    /**
     * class 中字段的信息
     */
    private static final Map<Class, List<FieldInfo>> CLASS_FIELD_INFO_CACHE = new ConcurrentReferenceHashMap<>();

    private static final Map<Class, List<List<String>>> CLASS_HEAD_CACHE_MAP = new ConcurrentReferenceHashMap<>();

    /**
     * 获取class中表头相关信息
     * @param clazz
     * @return
     */
    public List<FieldInfo> getClassFieldInfo(Class<?> clazz) {
        List<FieldInfo> fieldInfos = CLASS_FIELD_INFO_CACHE.get(clazz);
        if (fieldInfos != null) {
            return fieldInfos;
        }
        fieldInfos = new ArrayList<>();
        Field[] fields = clazz.getDeclaredFields();
        for (Field field : fields) {
            FieldInfo fieldInfo = new FieldInfo();
            fieldInfo.setField(field);
            fieldInfo.setFieldName(field.getName());
            fieldInfo.setValidationData(field.getAnnotation(ValidationData.class));
            ExcelProperty excelProperty = field.getAnnotation(ExcelProperty.class);
            if (excelProperty != null) {
                fieldInfo.setHeadNames(ListUtil.toList(excelProperty.value()));
                if (excelProperty.index() != -1) {
                    fieldInfo.setHeadColumnIndex(excelProperty.index());
                }
            }
            ExtendColumn extendColumn = field.getAnnotation(ExtendColumn.class);
            if (extendColumn != null) {
                fieldInfo.setExtendColumn(true);
            }
            fieldInfo.setHeadFontProperty(FontProperty.build(field.getAnnotation(HeadFontStyle.class)));
            fieldInfo.setHeadStyleProperty(StyleProperty.build(field.getAnnotation(HeadStyle.class)));
            fieldInfo.setContentFontProperty(FontProperty.build(field.getAnnotation(ContentFontStyle.class)));
            fieldInfo.setContentStyleProperty(StyleProperty.build(field.getAnnotation(ContentStyle.class)));
            fieldInfo.setColumnWidthProperty(ColumnWidthProperty.build(field.getAnnotation(ColumnWidth.class)));
            fieldInfo.setLoopMergeProperty(LoopMergeProperty.build(field.getAnnotation(ContentLoopMerge.class)));
            fieldInfos.add(fieldInfo);
        }
        CLASS_FIELD_INFO_CACHE.put(clazz, fieldInfos);
        return fieldInfos;
    }

    /**
     * 获取 class 中的下拉框信息
     * @param clazz
     * @return
     */
    public List<ValidationInfo> getValidationInfoInfos(Class<?> clazz) {
        List<FieldInfo> fieldInfos = getClassFieldInfo(clazz);
        if (CollUtil.isEmpty(fieldInfos)) {
            return new ArrayList<>();
        }
        return fieldInfos.stream().filter(fieldInfo -> Objects.nonNull(fieldInfo.getValidationData()))
                .map(fieldInfo -> {
                    ValidationData validationData = fieldInfo.getValidationData();
                    return ValidationInfo.buildColumnByField(fieldInfo.getFieldName(), validationData.rowNum(), validationData.options())
                            .setAsDicSheet(validationData.asDicSheet())
                            .setSheetName(validationData.sheetName())
                            .setDicTitle(validationData.dicTitle())
                            .setCheckDatavalidity(validationData.checkDatavalidity())
                            .setErrorBox(validationData.errorTitle(),validationData.errorMsg());
                }).collect(Collectors.toList());
    }

    /**
     * 根据字段名获class取字段信息
     */
    public Optional<FieldInfo> getFieldInfoByFieldName(Class<?> clazz, String fieldName) {
        if (fieldName == null) {
            return Optional.empty();
        }
        List<FieldInfo> fieldInfos = getClassFieldInfo(clazz);
        if (CollUtil.isEmpty(fieldInfos)) {
            return Optional.empty();
        }
        return fieldInfos.stream().filter(info -> fieldName.equals(info.getFieldName()))
                .findFirst();
    }

    /**
     * 获取 clazz 中表头信息
     * @return  外层list下标为行索引，内层list下标为列索引
     */
    public List<List<String>> getClassHeads(Class<?> clazz) {
        List<List<String>> result = CLASS_HEAD_CACHE_MAP.get(clazz);
        if (result != null) {
            return result;
        }
        Map<Integer, Field> filedMap = new HashMap<>();
        com.alibaba.excel.util.ClassUtils.declaredFields(clazz, filedMap, null, null, true, false, null);
        if (CollUtil.isEmpty(filedMap)) {
            return null;
        }
        result = new ArrayList<>();
        for (Map.Entry<Integer, Field> entry : filedMap.entrySet()) {
            Field field = entry.getValue();
            ExcelProperty excelProperty = field.getAnnotation(ExcelProperty.class);
            if (excelProperty == null) {
                List<String> rowHeads;
                if (result.size() == 0) {
                    rowHeads = new ArrayList<>();
                    result.add(rowHeads);
                } else {
                    rowHeads = result.get(0);
                }
                rowHeads.add(entry.getKey(), field.getName());
                continue;
            }
            String[] headArray = excelProperty.value();
            for (int rowIndex = 0; rowIndex < headArray.length; rowIndex++) {
                if (StrUtil.isBlank(headArray[rowIndex])) {
                    continue;
                }
                List<String> rowHeads;
                if (rowIndex >= result.size()) {
                    rowHeads = new ArrayList<>();
                    result.add(rowHeads);
                } else {
                    rowHeads = result.get(rowIndex);
                }
                rowHeads.add(entry.getKey(), headArray[rowIndex]);
            }
        }
        CLASS_HEAD_CACHE_MAP.put(clazz, result);
        return result;
    }

}

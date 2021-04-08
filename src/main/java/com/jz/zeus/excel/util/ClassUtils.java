package com.jz.zeus.excel.util;

import cn.hutool.core.collection.CollUtil;
import cn.hutool.core.util.ArrayUtil;
import cn.hutool.core.util.StrUtil;
import com.alibaba.excel.annotation.ExcelProperty;
import com.jz.zeus.excel.DropDownBoxInfo;
import com.jz.zeus.excel.FieldInfo;
import com.jz.zeus.excel.annotation.DropDownBox;
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
            fieldInfo.setFieldName(field.getName());
            ExcelProperty excelProperty = field.getAnnotation(ExcelProperty.class);
            if (excelProperty != null) {
                fieldInfo.setHeadName(excelProperty.value()[0]);
                if (excelProperty.index() != -1) {
                    fieldInfo.setHeadColumnIndex(excelProperty.index());
                }
            }
            DropDownBox dropDownBox = field.getAnnotation(DropDownBox.class);
            if (dropDownBox != null) {
                fieldInfo.setDropDownBoxOptions(dropDownBox.options());
                fieldInfo.setDropDownBoxRowNum(dropDownBox.rowNum());
            }
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
    public List<DropDownBoxInfo> getDropDownBoxInfos(Class<?> clazz) {
        List<FieldInfo> fieldInfos = getClassFieldInfo(clazz);
        if (CollUtil.isEmpty(fieldInfos)) {
            return new ArrayList<>();
        }
        return fieldInfos.stream().filter(fieldInfo -> (fieldInfo.getHeadColumnIndex() != null || StrUtil.isNotBlank(fieldInfo.getHeadName()))
                                    && ArrayUtil.isNotEmpty(fieldInfo.getDropDownBoxOptions()))
                .map(fieldInfo -> {
                    if (fieldInfo.getHeadColumnIndex() != null) {
                        return new DropDownBoxInfo(fieldInfo.getHeadColumnIndex(), fieldInfo.getDropDownBoxRowNum(), fieldInfo.getDropDownBoxOptions());
                    }
                    return new DropDownBoxInfo(fieldInfo.getHeadName(), fieldInfo.getDropDownBoxRowNum(), fieldInfo.getDropDownBoxOptions());
                }).collect(Collectors.toList());
    }

    public Optional<FieldInfo> getFieldInfoByFieldName(Class<?> clazz, String fieldName) {
        if (fieldName == null) {
            return Optional.empty();
        }
        List<FieldInfo> fieldInfos = CLASS_FIELD_INFO_CACHE.get(clazz);
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

package com.jz.zeus.excel;

import cn.hutool.core.collection.ListUtil;
import cn.hutool.core.lang.Assert;
import cn.hutool.core.text.CharSequenceUtil;
import cn.hutool.core.util.ArrayUtil;
import cn.hutool.core.util.RandomUtil;
import com.jz.zeus.excel.util.ClassUtils;
import lombok.AccessLevel;
import lombok.Getter;
import lombok.Setter;
import lombok.experimental.Accessors;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.Objects;

/**
 * @Author JZ
 * @Date 2021/3/23 18:27
 */
@Setter
@Getter
@Accessors(chain = true)
public class ValidationInfo {

    private static final Integer DEFAULT_ROW_NUM = 10000;

    private static final Integer DEFAULT_COLUMN_NUM = 100;

    @Setter(AccessLevel.NONE)
    private Integer rowIndex;

    @Setter(AccessLevel.NONE)
    private Integer rowNum;

    @Setter(AccessLevel.NONE)
    private Integer columnIndex;

    @Setter(AccessLevel.NONE)
    private Integer columnNum = 100;

    @Setter(AccessLevel.NONE)
    private String fieldName;

    @Setter(AccessLevel.NONE)
    private String headName;

    @Setter(AccessLevel.NONE)
    private List<String> options;

    /**
     * 是否校验数据是否为下拉框内容，若为false {@link #errorTitle} 和 {@link #errorMsg} 将无效
     */
    private boolean checkDatavalidity = true;

    /**
     * 错误信息box的标题
     */
    private String errorTitle;

    /**
     * 当输入的数据不是下拉框中数据时的提示信息
     */
    private String errorMsg;

    /**
     * sheet 名称
     */
    @Getter(AccessLevel.NONE)
    private String sheetName;

    /**
     * 为 true 不论选项条数是否超过设定值，该下拉信息都将生成到sheet中且不会被隐藏
     */
    private boolean asDicSheet;

    /**
     * 父级下拉框
     */
    @Setter(AccessLevel.NONE)
    private ValidationInfo parent;

    /**
     * 父级下拉框与子级下拉框的映射关系
     */
    @Setter(AccessLevel.NONE)
    private Map<String, List<String>> parentChildMap;

    private ValidationInfo() {}

    public static ValidationInfo buildCascadeByHead(String headName, ValidationInfo parent, Map<String, List<String>> parentChildMap) {
        return buildCascade(null, null, headName, null, parent, parentChildMap);
    }

    public static ValidationInfo buildCascadeByHead(String headName, Integer rowNum, ValidationInfo parent, Map<String, List<String>> parentChildMap) {
        return buildCascade(null, null, headName, rowNum, parent, parentChildMap);
    }

    public static ValidationInfo buildCascadeByField(String fieldName, ValidationInfo parent, Map<String, List<String>> parentChildMap) {
        return buildCascade(null, fieldName, null, null, parent, parentChildMap);
    }

    public static ValidationInfo buildCascadeByField(String fieldName, Integer rowNum, ValidationInfo parent, Map<String, List<String>> parentChildMap) {
        return buildCascade(null, fieldName, null, rowNum, parent, parentChildMap);
    }

    public static ValidationInfo buildCascadeByIndex(int columnIndex, ValidationInfo parent, Map<String, List<String>> parentChildMap) {
        return buildCascade(columnIndex, null, null, null, parent, parentChildMap);
    }

    public static ValidationInfo buildCascadeByIndex(int columnIndex, Integer rowNum, ValidationInfo parent, Map<String, List<String>> parentChildMap) {
        return buildCascade(columnIndex, null, null, rowNum, parent, parentChildMap);
    }

    private static ValidationInfo buildCascade(Integer columnIndex, String fieldName, String headName, Integer rowNum,
                                                     ValidationInfo parent, Map<String, List<String>> parentChildMap) {
        verifyColumn(columnIndex, fieldName, headName, null, rowNum);
        ValidationInfo validationInfo = new ValidationInfo();
        validationInfo.columnIndex = columnIndex;
        validationInfo.fieldName = fieldName;
        validationInfo.headName = headName;
        validationInfo.rowNum = (rowNum == null ? DEFAULT_ROW_NUM : rowNum);
        validationInfo.parent = parent;
        validationInfo.parentChildMap = parentChildMap;
        return validationInfo;
    }

    public static ValidationInfo buildColumnByIndex(int columnIndex, String... options) {
        return buildColumn(null, null, columnIndex, null, options);
    }

    public static ValidationInfo buildColumnByIndex(int columnIndex, List<String> options) {
        return buildColumn(null, null, columnIndex, null, options);
    }

    public static ValidationInfo buildColumnByIndex(int columnIndex, Integer rowNum, String... options) {
        return buildColumn(null, null, columnIndex, rowNum, options);
    }

    public static ValidationInfo buildColumnByIndex(int columnIndex, Integer rowNum, List<String> options) {
        return buildColumn(null, null, columnIndex, rowNum, options);
    }

    public static ValidationInfo buildColumnByHead(String headName, String... options) {
        return buildColumn(null, headName, null, null, options);
    }
    public static ValidationInfo buildColumnByHead(String headName, List<String> options) {
        return buildColumn(null, headName, null, null, options);
    }

    public static ValidationInfo buildColumnByHead(String headName, Integer rowNum, String... options) {
        return buildColumn(null, headName, null, rowNum, options);
    }

    public static ValidationInfo buildColumnByHead(String headName, Integer rowNum, List<String> options) {
        return buildColumn(null, headName, null, rowNum, options);
    }

    public static ValidationInfo buildColumnByField(String fieldName, String... options) {
        return buildColumn(fieldName, null, null, null, options);
    }

    public static ValidationInfo buildColumnByField(String fieldName, List<String> options) {
        return buildColumn(fieldName, null, null, null, options);
    }

    public static ValidationInfo buildColumnByField(String fieldName, Integer rowNum, String... options) {
        return buildColumn(fieldName, null, null, rowNum, options);
    }

    public static ValidationInfo buildColumnByField(String fieldName, Integer rowNum, List<String> options) {
        return buildColumn(fieldName, null, null, rowNum, options);
    }

    public static ValidationInfo buildColumnByField(Class<?> clazz, String fieldName, String... options) {
        if (ClassUtils.getFieldInfoByFieldName(clazz, fieldName).isPresent()) {
            return buildColumn(fieldName, null, null, null, options);
        }
        return null;
    }

    public static ValidationInfo buildColumnByField(Class<?> clazz, String fieldName, List<String> options) {
        if (ClassUtils.getFieldInfoByFieldName(clazz, fieldName).isPresent()) {
            return buildColumn(fieldName, null, null, null, options);
        }
        return null;
    }

    public static ValidationInfo buildColumnByField(Class<?> clazz, String fieldName, Integer rowNum, String... options) {
        if (ClassUtils.getFieldInfoByFieldName(clazz, fieldName).isPresent()) {
            return buildColumn(fieldName, null, null, rowNum, options);
        }
        return null;
    }

    public static ValidationInfo buildColumnByField(Class<?> clazz, String fieldName, Integer rowNum, List<String> options) {
        if (ClassUtils.getFieldInfoByFieldName(clazz, fieldName).isPresent()) {
            return buildColumn(fieldName, null, null, rowNum, options);
        }
        return null;
    }

    private static ValidationInfo buildColumn(String fieldName, String headName, Integer columnIndex, Integer rowNum, String... options) {
        if (ArrayUtil.isEmpty(options)) {
            return buildColumn(fieldName, headName, columnIndex, rowNum, new ArrayList<>(0));
        }
        return buildColumn(fieldName, headName, columnIndex, rowNum, ListUtil.toList(options));
    }

    private static ValidationInfo buildColumn(String fieldName, String headName, Integer columnIndex, Integer rowNum, List<String> options) {
        verifyColumn(columnIndex, fieldName, headName, null, rowNum);
        ValidationInfo validationInfo = new ValidationInfo();
        validationInfo.columnIndex = columnIndex;
        validationInfo.fieldName = fieldName;
        validationInfo.headName = headName;
        validationInfo.rowNum = (rowNum == null ? DEFAULT_ROW_NUM : rowNum);
        validationInfo.options = options;
        return validationInfo;
    }

    /**
     * 构建具有相同选项的多列下拉框
     * @param names                 表头 或 属性名
     * @param isField               true names 表示属性名、false names 表示表头
     * @param options               选项内容
     */
    public static List<ValidationInfo> buildCommonOption(String[] names, boolean isField, String... options) {
        if (ArrayUtil.isEmpty(names)) {
            return new ArrayList<>(0);
        }
        List<ValidationInfo> validationInfoList = new ArrayList<>();
        for(String name : names) {
            if (isField) {
                validationInfoList.add(ValidationInfo.buildColumnByField(name, options));
            } else {
                validationInfoList.add(ValidationInfo.buildColumnByHead(name, options));
            }
        }
        return validationInfoList;
    }

    /**
     * 构建具有相同选项的多列下拉框
     * @param columnIndexs          列索引
     * @param options               选项内容
     */
    public static List<ValidationInfo> buildCommonOption(Integer[] columnIndexs, String... options) {
        if (ArrayUtil.isEmpty(columnIndexs) || ArrayUtil.isEmpty(options)) {
            return new ArrayList<>(0);
        }
        List<ValidationInfo> validationInfoList = new ArrayList<>(columnIndexs.length);
        for(int i = 0; i < columnIndexs.length; ++i) {
            validationInfoList.add(ValidationInfo.buildColumnByIndex(columnIndexs[i], options));
        }
        return validationInfoList;
    }

    /**
     * 构建某一行的下拉框
     */
    public static ValidationInfo buildRow(Integer rowIndex, String... options) {
        return buildRow(rowIndex, null, options);
    }

    /**
     * 构建某一行的下拉框
     */
    public static ValidationInfo buildRow(Integer rowIndex, Integer columnNum, String... options) {
        Assert.isTrue(rowIndex != null, "RowIndex can't be null");
        if (rowIndex != null) {
            Assert.isTrue(rowIndex >= 0, "RowIndex has to be greater than or equal to 0");
        }
        if (columnNum != null) {
            Assert.isTrue(columnNum >= 0, "ColumnNum has to be greater than or equal to 0");
        }
        ValidationInfo info = new ValidationInfo();
        info.rowIndex = rowIndex;
        info.columnNum = (columnNum == null ? DEFAULT_COLUMN_NUM : columnNum);
        info.options = ListUtil.toList(options);
        return info;
    }

    /**
     * 构建某个单元格的下拉框
     */
    public static ValidationInfo buildPrecise(Integer rowIndex, Integer columnIndex, String... options) {
        Assert.isTrue(rowIndex != null && columnIndex != null, "RowIndex and ColumnIndex can't both be null");
        Assert.isTrue(rowIndex >= 0, "RowIndex has to be greater than or equal to 0");
        Assert.isTrue(columnIndex >= 0, "ColumnNum has to be greater than or equal to 0");
        ValidationInfo info = new ValidationInfo();
        info.rowIndex = rowIndex;
        info.columnIndex = columnIndex;
        info.options = ListUtil.toList(options);
        return info;
    }

    /**
     * 构建某个单元格的下拉框
     */
    public static ValidationInfo buildPrecise(Integer rowIndex, String headName, String... options) {
        Assert.isTrue(rowIndex != null && CharSequenceUtil.isNotBlank(headName), "RowIndex and HeadName can't both be empty");
        Assert.isTrue(rowIndex >= 0, "RowIndex has to be greater than or equal to 0");
        ValidationInfo info = new ValidationInfo();
        info.rowIndex = rowIndex;
        info.headName = headName;
        info.options = ListUtil.toList(options);
        return info;
    }

    /**
     * 构建字典表
     */
    public static ValidationInfo buildDictionaryTable(String sheetName, List<String> options) {
        ValidationInfo info = new ValidationInfo();
        info.asDicSheet = true;
        info.sheetName = sheetName;
        info.options = options;
        return info;
    }

    /**
     * 构建字典表
     */
    public static ValidationInfo buildDictionaryTable(String sheetName, String... options) {
        ValidationInfo info = new ValidationInfo();
        info.asDicSheet = true;
        info.sheetName = sheetName;
        info.options = ListUtil.toList(options);
        return info;
    }


    public ValidationInfo asDicSheet(String sheetName) {
        this.asDicSheet = true;
        this.sheetName = sheetName;
        return this;
    }

    public String getSheetName() {
        if (CharSequenceUtil.isBlank(sheetName)) {
            sheetName = "dic_" + RandomUtil.randomString(16);
        }
        return sheetName;
    }

    public ValidationInfo setErrorBox(String errorTitle, String errorMsg) {
        this.errorTitle = errorTitle;
        this.errorMsg = errorMsg;
        return this;
    }

    @Override
    public boolean equals(Object o) {
        if (this == o) {
            return true;
        }
        if (!(o instanceof ValidationInfo)) {
            return false;
        }
        ValidationInfo that = (ValidationInfo) o;
        if (Objects.equals(rowIndex, that.rowIndex)) {
            if (rowIndex != null && Objects.equals(that.columnIndex, columnIndex)) {
                return true;
            } else if (CharSequenceUtil.isNotBlank(fieldName) && Objects.equals(that.fieldName, fieldName)) {
                return true;
            } else if (CharSequenceUtil.isNotBlank(headName) && Objects.equals(that.headName, headName)) {
                return true;
            }
        }
        return false;
    }

    @Override
    public int hashCode() {
        return Objects.hash(rowIndex, fieldName, rowNum, columnIndex, columnNum, headName, options);
    }

    private static void verifyColumn(Integer columnIndex, String fieldName, String headName, Integer rowIndex, Integer rowNum) {
        Assert.isFalse(columnIndex == null && CharSequenceUtil.isBlank(fieldName) && CharSequenceUtil.isBlank(headName),
                "ColumnIndex and FieldName and HeadName not all empty");
        if (rowIndex != null) {
            Assert.isTrue(rowIndex >= 0, "RowIndex has to be greater than or equal to 0");
        }
        if (rowNum != null) {
            Assert.isTrue(rowNum >= 0, "RowNum has to be greater than or equal to 0");
        }
        if (columnIndex != null) {
            Assert.isTrue(columnIndex >= 0, "ColumnIndex has to be greater than or equal to 0");
        }
    }

}

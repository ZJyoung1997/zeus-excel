package com.jz.zeus.excel.write.handler;

import cn.hutool.core.collection.CollUtil;
import com.alibaba.excel.write.metadata.holder.WriteSheetHolder;
import com.alibaba.excel.write.metadata.holder.WriteTableHolder;
import com.alibaba.excel.write.metadata.holder.WriteWorkbookHolder;
import com.jz.zeus.excel.CellErrorInfo;
import com.jz.zeus.excel.context.ExcelContext;
import com.jz.zeus.excel.util.ExcelUtils;
import com.jz.zeus.excel.write.helper.WriteSheetHelper;
import lombok.Setter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.stream.Collectors;

/**
 * @Author JZ
 * @Date 2021/3/26 10:24
 */
public class ErrorInfoHandler extends AbstractRowWriteHandler {

    private ExcelContext excelContext;

    private WriteSheetHelper writeSheetHelper;

    private Integer headRowNum;

    @Setter
    private String commentRowPrefix = "- ";

    @Setter
    private String commentRowSuffix = "";

    /**
     * 下一个需要处理的行索引
     */
    private int nextRowIndex;

    private Map<Integer, List<CellErrorInfo>> rowErrorInfoMap;

    private Sheet currentSheet;

    public ErrorInfoHandler(ExcelContext excelContext, List<CellErrorInfo> errorInfoList) {
        this(excelContext, null, errorInfoList);
    }

    /**
     * @param headRowNum     仅对用表头名作为列坐标时有影响，有误会导致加载不到对应表头，导致对应单元格无法添加错误信息
     * @param errorInfoList
     */
    public ErrorInfoHandler(ExcelContext excelContext, Integer headRowNum, List<CellErrorInfo> errorInfoList) {
        this.excelContext = excelContext;
        this.headRowNum = headRowNum;
        if (CollUtil.isNotEmpty(errorInfoList)) {
            rowErrorInfoMap = errorInfoList.stream()
                    .collect(Collectors.groupingBy(CellErrorInfo::getRowIndex));
        }
    }

    @Override
    public void afterRowDispose(WriteSheetHolder writeSheetHolder, WriteTableHolder writeTableHolder, Row row, Integer relativeRowIndex, Boolean isHead) {
        if (writeSheetHelper == null) {
            return;
        }
        createComment(currentSheet, row.getRowNum());
    }

    @Override
    public void afterSheetCreate(WriteWorkbookHolder writeWorkbookHolder, WriteSheetHolder writeSheetHolder) {
        if (CollUtil.isEmpty(rowErrorInfoMap)) {
            return;
        }
        currentSheet = getSheet(writeWorkbookHolder, writeSheetHolder);
        writeSheetHelper = new WriteSheetHelper(excelContext, writeSheetHolder, headRowNum);
        Boolean needHead = writeSheetHolder.getNeedHead();
        if (Boolean.FALSE.equals(needHead) && CollUtil.isEmpty(excelContext.getSheetData())) {
            Integer maxRowIndex = rowErrorInfoMap.keySet().stream().filter(Objects::nonNull).max(Integer::compare)
                    .orElse(-1);
            createComment(currentSheet, maxRowIndex);
        }
    }

    private Sheet getSheet(WriteWorkbookHolder writeWorkbookHolder, WriteSheetHolder writeSheetHolder) {
        Workbook workbook = writeWorkbookHolder.getWorkbook();
        Sheet sheet;
        if (workbook instanceof SXSSFWorkbook) {
            if (writeSheetHolder.getSheetNo() != null) {
                sheet = ((SXSSFWorkbook) workbook).getXSSFWorkbook().getSheetAt(writeSheetHolder.getSheetNo());
            } else {
                sheet = ((SXSSFWorkbook) workbook).getXSSFWorkbook().getSheet(writeSheetHolder.getSheetName());
            }
        } else {
            sheet = writeSheetHolder.getSheet();
        }
        return sheet;
    }

    private void createComment(Sheet sheet, Integer currentRowIndex) {
        for (int i = nextRowIndex; i <= currentRowIndex; i++) {
            List<CellErrorInfo> rowErrorInfos = rowErrorInfoMap.get(i);
            if (CollUtil.isEmpty(rowErrorInfos)) {
                continue;
            }
            int finalI = i;
            rowErrorInfos.stream()
                    .collect(Collectors.groupingBy(e -> getColumnIndex(e)))
                    .forEach((columnIndex, errorInfos) -> {
                        if (columnIndex == -1) {
                            return;
                        }
                        List<String> errorMsgs = errorInfos.stream()
                                .flatMap(e -> e.getErrorMsgs().stream())
                                .collect(Collectors.toList());
                        ExcelUtils.setCommentErrorInfo(sheet, finalI, columnIndex, commentRowPrefix, commentRowSuffix,
                                errorMsgs.toArray(new String[0]));
                    });
        }
        nextRowIndex = currentRowIndex + 1;
    }

    private Integer getColumnIndex(CellErrorInfo errorInfo) {
        Integer columnIndex;
        if ((columnIndex = errorInfo.getColumnIndex()) != null) {
            return columnIndex;
        } else if ((columnIndex = writeSheetHelper.getFieldColumnIndex(errorInfo.getFieldName())) != null) {
            return columnIndex;
        } else if ((columnIndex = writeSheetHelper.getHeadColumnIndex(errorInfo.getHeadName())) != null) {
            return columnIndex;
        }
        return -1;
    }

}

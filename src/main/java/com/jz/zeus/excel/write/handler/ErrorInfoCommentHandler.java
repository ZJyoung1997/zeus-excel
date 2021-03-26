package com.jz.zeus.excel.write.handler;

import cn.hutool.core.collection.CollUtil;
import cn.hutool.core.util.StrUtil;
import com.alibaba.excel.write.handler.AbstractSheetWriteHandler;
import com.alibaba.excel.write.metadata.holder.WriteSheetHolder;
import com.alibaba.excel.write.metadata.holder.WriteWorkbookHolder;
import com.jz.zeus.excel.CellErrorInfo;
import lombok.Setter;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;

import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;

/**
 * @Author JZ
 * @Date 2021/3/26 10:24
 */
public class ErrorInfoCommentHandler extends AbstractSheetWriteHandler {

    /**
     * 表头行数
     */
    private Integer headRowNum;

    @Setter
    private String commentRowPrefix = "- ";

    @Setter
    private String commentRowSuffix = "";

    private Map<Integer, List<CellErrorInfo>> rowErrorInfoMap;

    /**
     * 存放表头与其列索引的信息
     * key 表头、value 列索引
     */
    private Map<String, Integer> headNameIndexMap;

    public ErrorInfoCommentHandler(List<CellErrorInfo> errorInfoList) {
        this(1, errorInfoList);
    }

    public ErrorInfoCommentHandler(Integer headRowNum, List<CellErrorInfo> errorInfoList) {
        if (headRowNum != null) {
            this.headRowNum = headRowNum;
        }
        if (CollUtil.isNotEmpty(errorInfoList)) {
            this.rowErrorInfoMap = errorInfoList.stream()
                    .collect(Collectors.groupingBy(CellErrorInfo::getRowIndex));
        }
    }

    @Override
    public void afterSheetCreate(WriteWorkbookHolder writeWorkbookHolder, WriteSheetHolder writeSheetHolder) {
        if (CollUtil.isEmpty(rowErrorInfoMap)) {
            return;
        }
        Sheet sheet = writeSheetHolder.getCachedSheet();
        initHeadNameIndexMap(sheet);
        rowErrorInfoMap.forEach((rowIndex, errorInfos) -> {
            Row row = sheet.getRow(rowIndex);
            if (row == null) {
                return;
            }
            errorInfos.forEach(errorInfo -> {
                setErrorInfo(sheet, row, errorInfo);
            });
        });
    }

    private void setErrorInfo(Sheet sheet, Row row, CellErrorInfo errorInfo) {
        Integer columnIndex = errorInfo.getColumnIndex();
        if (columnIndex == null) {
            columnIndex = headNameIndexMap.get(errorInfo.getHeadName());
        }
        Cell cell = row.getCell(columnIndex);
        if (cell == null) {
            return;
        }
        CellStyle cellStyle = sheet.getWorkbook().createCellStyle();
        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        cellStyle.setFillForegroundColor(IndexedColors.RED.index);
        cell.setCellStyle(cellStyle);

        Drawing<?> drawing = sheet.createDrawingPatriarch();
        Comment comment = drawing.createCellComment(new XSSFClientAnchor(0, 0, 0, 0, columnIndex, errorInfo.getRowIndex(), columnIndex+2, errorInfo.getRowIndex()+2));
        comment.setString(new XSSFRichTextString(CollUtil.join(errorInfo.getErrorMsgs(), "\n", commentRowPrefix, commentRowSuffix)));
        cell.setCellComment(comment);
    }

    /**
     * 初始化headNameIndexMap
     * @param sheet
     */
    private void initHeadNameIndexMap(Sheet sheet) {
        if (CollUtil.isEmpty(headNameIndexMap)) {
            headNameIndexMap = new HashMap<>();
        }
        for (int i = 0; i < headRowNum; i++) {
            Row headRow = sheet.getRow(i);
            int cellNum = headRow.getLastCellNum();
            for (int j = 0; j < cellNum; j++) {
                Cell cell = headRow.getCell(j);
                if (cell == null || StrUtil.isBlank(cell.getStringCellValue())) {
                    continue;
                }
                headNameIndexMap.put(cell.getStringCellValue(), j);
            }
        }
    }

}

package com.jz.zeus.excel.write.handler;

import cn.hutool.core.collection.CollUtil;
import com.alibaba.excel.write.metadata.holder.WriteSheetHolder;
import com.alibaba.excel.write.metadata.holder.WriteWorkbookHolder;
import com.jz.zeus.excel.CellErrorInfo;
import com.jz.zeus.excel.util.ExcelUtils;
import lombok.Setter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;

/**
 * @Author JZ
 * @Date 2021/3/26 10:24
 */
public class ErrorInfoCommentHandler extends AbstractZeusSheetWriteHandler {

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
        this(null, errorInfoList);
    }

    /**
     * @param headRowNum     仅对用表头名作为列坐标时有影响，有误会导致加载不到对应表头，导致对应单元格无法添加错误信息
     * @param errorInfoList
     */
    public ErrorInfoCommentHandler(Integer headRowNum, List<CellErrorInfo> errorInfoList) {
        super(headRowNum);
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
        initHeadNameIndexMap(writeSheetHolder);
        rowErrorInfoMap.forEach((rowIndex, errorInfos) -> {
            Row row = sheet.getRow(rowIndex);
            if (row == null) {
                return;
            }
            errorInfos.forEach(errorInfo -> {
                Integer columnIndex = errorInfo.getColumnIndex();
                if (columnIndex == null) {
                    columnIndex = headNameIndexMap.get(errorInfo.getHeadName());
                }
                ExcelUtils.setCommentErrorInfo(sheet, rowIndex, columnIndex, errorInfo.getErrorMsgs().toArray(new String[0]));
            });
        });
    }

}

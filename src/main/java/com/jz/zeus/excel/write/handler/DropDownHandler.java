package com.jz.zeus.excel.write.handler;

import cn.hutool.core.util.StrUtil;
import com.alibaba.excel.metadata.CellData;
import com.alibaba.excel.metadata.Head;
import com.alibaba.excel.write.handler.AbstractCellWriteHandler;
import com.alibaba.excel.write.handler.SheetWriteHandler;
import com.alibaba.excel.write.metadata.holder.WriteSheetHolder;
import com.alibaba.excel.write.metadata.holder.WriteTableHolder;
import com.alibaba.excel.write.metadata.holder.WriteWorkbookHolder;
import com.jz.zeus.excel.DropDownBoxInfo;
import lombok.NoArgsConstructor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataValidationHelper;
import org.apache.poi.ss.usermodel.Sheet;

import java.util.ArrayList;
import java.util.List;

/**
 * @Author JZ
 * @Date 2021/3/23 17:55
 */
@NoArgsConstructor
public class DropDownHandler extends AbstractCellWriteHandler {

    private List<DropDownBoxInfo> columnBoxInfoList = new ArrayList<>();

    private List<DropDownBoxInfo> rowBoxInfoList = new ArrayList<>();

    private List<DropDownBoxInfo> otherBoxInfoList = new ArrayList<>();

    public DropDownHandler(List<DropDownBoxInfo> dropDownBoxInfos) {
        dropDownBoxInfos.forEach(info -> addDropDownBoxInfo(info));
    }

    @Override
    public void afterCellDispose(WriteSheetHolder writeSheetHolder, WriteTableHolder writeTableHolder, List<CellData> cellDataList, Cell cell, Head head, Integer relativeRowIndex, Boolean isHead) {
        if (Boolean.TRUE.equals(isHead)) {
            columnBoxInfoList.forEach(info -> {
                if ((info.getColumnIndex() != null && info.getColumnIndex().intValue() == cell.getColumnIndex())
            });
            return;
        }

    }

    public void addDropDownBoxInfo(DropDownBoxInfo boxInfo) {
        if (boxInfo.getRowIndex() == null
                && (boxInfo.getColumnIndex() != null || StrUtil.isNotBlank(boxInfo.getHeadName()))) {
            columnBoxInfoList.add(boxInfo);
        } else if (boxInfo.getRowIndex() != null && boxInfo.getColumnIndex() == null
                && StrUtil.isNotBlank(boxInfo.getHeadName())) {
            rowBoxInfoList.add(boxInfo);
        } else {
            otherBoxInfoList.add(boxInfo);
        }
    }

}

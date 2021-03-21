package com.jz.zeus.excel;

import cn.hutool.core.collection.CollectionUtil;
import com.alibaba.excel.support.ExcelTypeEnum;
import lombok.Data;

import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

/**
 * @author:JZ
 * @date:2021/3/18
 */
@Data
public class ExcelInfo {

    /**
     * excel 名称
     */
    private String name;

    /**
     * Excel的文件格式
     */
    private ExcelTypeEnum suffix = ExcelTypeEnum.XLSX;

    /**
     * Excel密码
     */
    private String password;

    /**
     * 是否在内存中操作
     */
    private boolean inMemory;

    /**
     * sheet信息
     */
    private List<SheetInfo> sheetInfos;

    public void addSheetInfo(SheetInfo sheetInfo) {
        if (CollectionUtil.isEmpty(sheetInfos)) {
            sheetInfos = new ArrayList<SheetInfo>();
        }
        sheetInfos.add(sheetInfo);
    }

    public void removeSheetInfo(String sheetName) {
        if (CollectionUtil.isEmpty(sheetInfos)) {
            return;
        }
        Iterator<SheetInfo> iterator = sheetInfos.iterator();
        while (iterator.hasNext()) {
            SheetInfo sheetInfo = iterator.next();
            if (sheetName.equals(sheetInfo.getName())) {
                iterator.remove();
                return;
            }
        }
    }

}

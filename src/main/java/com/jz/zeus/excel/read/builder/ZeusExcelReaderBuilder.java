package com.jz.zeus.excel.read.builder;

import cn.hutool.core.util.ReflectUtil;
import com.alibaba.excel.ExcelReader;
import com.alibaba.excel.analysis.ExcelAnalyserImpl;
import com.alibaba.excel.cache.ReadCache;
import com.alibaba.excel.cache.selector.ReadCacheSelector;
import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.context.xlsx.DefaultXlsxReadContext;
import com.alibaba.excel.enums.CellExtraTypeEnum;
import com.alibaba.excel.event.AnalysisEventListener;
import com.alibaba.excel.event.SyncReadListener;
import com.alibaba.excel.read.builder.ExcelReaderBuilder;
import com.alibaba.excel.read.builder.ExcelReaderSheetBuilder;
import com.alibaba.excel.support.ExcelTypeEnum;
import com.jz.zeus.excel.read.ZeusAnalysisEventProcessor;
import com.jz.zeus.excel.read.listener.ModelBuildEventListener;

import javax.xml.parsers.SAXParserFactory;
import java.io.File;
import java.io.InputStream;
import java.util.List;

/**
 * @Author JZ
 * @Date 2021/6/25 15:06
 */
public class ZeusExcelReaderBuilder {

    private ExcelReaderBuilder excelReaderBuilder;

    private boolean useEasyExcelDefaultListener;

    private boolean useZeusDefaultListener;

    public ZeusExcelReaderBuilder() {
        this.useZeusDefaultListener = true;
        excelReaderBuilder = new ExcelReaderBuilder();
    }

    public ZeusExcelReaderBuilder excelType(ExcelTypeEnum excelType) {
        excelReaderBuilder.excelType(excelType);
        return this;
    }

    public ZeusExcelReaderBuilder useEasyExcelDefaultListener(boolean useEasyExcelDefaultListener) {
        this.useEasyExcelDefaultListener = useEasyExcelDefaultListener;
        return this;
    }

    public ZeusExcelReaderBuilder file(InputStream inputStream) {
        excelReaderBuilder.file(inputStream);
        return this;
    }

    /**
     * Read file
     * <p>
     * If 'inputStream' and 'file' all not empty,file first
     */
    public ZeusExcelReaderBuilder file(File file) {
        excelReaderBuilder.file(file);
        return this;
    }

    /**
     * Read file
     * <p>
     * If 'inputStream' and 'file' all not empty,file first
     */
    public ZeusExcelReaderBuilder file(String pathName) {
        return file(new File(pathName));
    }

    /**
     * Mandatory use 'inputStream' .Default is false.
     * <p>
     * if false,Will transfer 'inputStream' to temporary files to improve efficiency
     */
    public ZeusExcelReaderBuilder mandatoryUseInputStream(Boolean mandatoryUseInputStream) {
        excelReaderBuilder.mandatoryUseInputStream(mandatoryUseInputStream);
        return this;
    }

    /**
     * Default true
     *
     * @param autoCloseStream
     * @return
     */
    public ZeusExcelReaderBuilder autoCloseStream(Boolean autoCloseStream) {
        excelReaderBuilder.autoCloseStream(autoCloseStream);
        return this;
    }

    /**
     * Ignore empty rows.Default is true.
     *
     * @param ignoreEmptyRow
     * @return
     */
    public ZeusExcelReaderBuilder ignoreEmptyRow(Boolean ignoreEmptyRow) {
        excelReaderBuilder.ignoreEmptyRow(ignoreEmptyRow);
        return this;
    }

    /**
     * This object can be read in the Listener {@link AnalysisEventListener#invoke(Object, AnalysisContext)}
     * {@link AnalysisContext#getCustom()}
     *
     * @param customObject
     * @return
     */
    public ZeusExcelReaderBuilder customObject(Object customObject) {
        excelReaderBuilder.customObject(customObject);
        return this;
    }

    /**
     * A cache that stores temp data to save memory.
     *
     * @param readCache
     * @return
     */
    public ZeusExcelReaderBuilder readCache(ReadCache readCache) {
        excelReaderBuilder.readCache(readCache);
        return this;
    }

    /**
     * Select the cache.Default use {@link com.alibaba.excel.cache.selector.SimpleReadCacheSelector}
     *
     * @param readCacheSelector
     * @return
     */
    public ZeusExcelReaderBuilder readCacheSelector(ReadCacheSelector readCacheSelector) {
        excelReaderBuilder.readCacheSelector(readCacheSelector);
        return this;
    }

    /**
     * Whether the encryption
     *
     * @param password
     * @return
     */
    public ZeusExcelReaderBuilder password(String password) {
        excelReaderBuilder.password(password);
        return this;
    }

    /**
     * SAXParserFactory used when reading xlsx.
     * <p>
     * The default will automatically find.
     * <p>
     * Please pass in the name of a class ,like : "com.sun.org.apache.xerces.internal.jaxp.SAXParserFactoryImpl"
     *
     * @see SAXParserFactory#newInstance()
     * @see SAXParserFactory#newInstance(String, ClassLoader)
     * @param xlsxSAXParserFactoryName
     * @return
     */
    public ZeusExcelReaderBuilder xlsxSAXParserFactoryName(String xlsxSAXParserFactoryName) {
        excelReaderBuilder.xlsxSAXParserFactoryName(xlsxSAXParserFactoryName);
        return this;
    }

    /**
     * Read some extra information, not by default
     *
     * @param extraType
     *            extra information type
     * @return
     */
    public ZeusExcelReaderBuilder extraRead(CellExtraTypeEnum extraType) {
        excelReaderBuilder.extraRead(extraType);
        return this;
    }

    public ExcelReader build() {
        excelReaderBuilder.useDefaultListener(useEasyExcelDefaultListener);
        if (useZeusDefaultListener) {
            excelReaderBuilder.registerReadListener(new ModelBuildEventListener());
        }
        ExcelReader excelReader = excelReaderBuilder.build();
        ExcelAnalyserImpl excelAnalyser = (ExcelAnalyserImpl) ReflectUtil.getFieldValue(excelReader, "excelAnalyser");
        DefaultXlsxReadContext xlsxReadContext = (DefaultXlsxReadContext) ReflectUtil.getFieldValue(excelAnalyser, "analysisContext");
        ReflectUtil.setFieldValue(xlsxReadContext, "analysisEventProcessor", new ZeusAnalysisEventProcessor());
        return excelReader;
    }

    public void doReadAll() {
        ExcelReader excelReader = build();
        excelReader.readAll();
        excelReader.finish();
    }

    /**
     * Synchronous reads return results
     *
     * @return
     */
    public <T> List<T> doReadAllSync() {
        SyncReadListener syncReadListener = new SyncReadListener();
        excelReaderBuilder.registerReadListener(syncReadListener);
        ExcelReader excelReader = build();
        excelReader.readAll();
        excelReader.finish();
        return (List<T>)syncReadListener.getList();
    }

    public ExcelReaderSheetBuilder sheet() {
        return sheet(null, null);
    }

    public ExcelReaderSheetBuilder sheet(Integer sheetNo) {
        return sheet(sheetNo, null);
    }

    public ExcelReaderSheetBuilder sheet(String sheetName) {
        return sheet(null, sheetName);
    }

    public ExcelReaderSheetBuilder sheet(Integer sheetNo, String sheetName) {
        ExcelReaderSheetBuilder excelReaderSheetBuilder = new ExcelReaderSheetBuilder(build());
        if (sheetNo != null) {
            excelReaderSheetBuilder.sheetNo(sheetNo);
        }
        if (sheetName != null) {
            excelReaderSheetBuilder.sheetName(sheetName);
        }
        return excelReaderSheetBuilder;
    }

}

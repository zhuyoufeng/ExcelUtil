package com.sargeraswang.util.ExcelUtil;

import java.util.Collection;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

/**
 * 用于汇出多个sheet的Vo The <code>ExcelSheet</code>
 * 
 * @author sargeras.wang
 * @version 1.0, Created at 2013年10月25日
 */
public class ExcelSheet<T> {
    private String sheetName;
    private LinkedHashMap<String,String> mapHeaders;
    private List<ExcelHeaderCell> definedHeaders;
    private int headerSize;
    private Collection<T> dataset;

    /**
     * @return the sheetName
     */
    public String getSheetName() {
        return sheetName;
    }

    /**
     * Excel页签名称
     * 
     * @param sheetName
     *            the sheetName to set
     */
    public void setSheetName(String sheetName) {
        this.sheetName = sheetName;
    }

    public LinkedHashMap<String, String> getMapHeaders() {
        return mapHeaders;
    }

    public void setMapHeaders(LinkedHashMap<String, String> mapHeaders) {
        this.mapHeaders = mapHeaders;
    }

    public List<ExcelHeaderCell> getDefinedHeaders() {
        return definedHeaders;
    }

    public void setDefinedHeaders(List<ExcelHeaderCell> definedHeaders) {
        this.definedHeaders = definedHeaders;
    }

    /**
     * Excel数据集合
     * 
     * @return the dataset
     */
    public Collection<T> getDataset() {
        return dataset;
    }

    /**
     * @param dataset
     *            the dataset to set
     */
    public void setDataset(Collection<T> dataset) {
        this.dataset = dataset;
    }

    public int getHeaderSize() {
        return headerSize;
    }

    public void setHeaderSize(int headerSize) {
        this.headerSize = headerSize;
    }
}

package com.itdl.excel.listener;

import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.event.AnalysisEventListener;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class ExcelDataListener<T> extends AnalysisEventListener<T> {

    private List<T> dataList = new ArrayList<>();
    private Map headMap = new HashMap();


    @Override
    public void invoke(T t, AnalysisContext analysisContext) {
        dataList.add(t);
    }

    @Override
    public void doAfterAllAnalysed(AnalysisContext analysisContext) {

    }

    public List<T> getDataList() {
        return this.dataList;
    }

    @Override
    public void invokeHeadMap(Map<Integer, String> headMap, AnalysisContext context) {
        this.headMap = headMap;
    }

    public Map getHeadMap() {
        return headMap;
    }
}

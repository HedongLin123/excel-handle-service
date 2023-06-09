package com.itdl.excel.listener;

import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.event.AnalysisEventListener;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * @Description 数据监听
 * @Author IT动力
 * @Date 2023/04/20 11:39
 */
public class ExcelDataListener<T> extends AnalysisEventListener<T> {

    // 数据列表
    private List<T> dataList = new ArrayList<>();
    // 标题列表
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

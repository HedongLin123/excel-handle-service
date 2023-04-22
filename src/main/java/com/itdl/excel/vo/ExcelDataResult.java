package com.itdl.excel.vo;

import java.io.Serializable;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * @Description 读物Excel结果数据
 * @Author IT动力
 * @Date 2023/04/20 15:18
 */
public class ExcelDataResult implements Serializable {

    // 标题数据
    private Map head = new HashMap();

    // 内容数据
    private List<Map<Integer, String>> datas = new ArrayList<>();

    // sheet名称
    private String sheetName;

    public ExcelDataResult(String sheetName, Map head, List<Map<Integer, String>> datas) {
        this.head = head;
        this.sheetName = sheetName;
        this.datas = datas;
    }

    public Map getHead() {
        return head;
    }

    public void setHead(Map head) {
        this.head = head;
    }

    public List<Map<Integer, String>> getDatas() {
        return datas;
    }

    public void setDatas(List<Map<Integer, String>> datas) {
        this.datas = datas;
    }

    public String getSheetName() {
        return sheetName;
    }

    public void setSheetName(String sheetName) {
        this.sheetName = sheetName;
    }
}

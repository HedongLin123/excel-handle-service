package com.itdl.excel.core;

import cn.hutool.core.collection.CollUtil;
import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelReader;
import com.alibaba.excel.read.metadata.ReadSheet;
import com.alibaba.excel.write.metadata.style.WriteCellStyle;
import com.alibaba.excel.write.metadata.style.WriteFont;
import com.alibaba.excel.write.style.HorizontalCellStyleStrategy;
import com.alibaba.excel.write.style.column.SimpleColumnWidthStyleStrategy;
import com.itdl.excel.listener.ExcelDataListener;
import com.itdl.excel.util.ExcelFillCellMergePrevColUtils;
import com.itdl.excel.vo.ExcelDataResult;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.VerticalAlignment;

import java.io.File;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;
import java.util.Map;

@Slf4j
public class ExcelParse {

    public static List<ExcelDataResult> simpleRead(InputStream inputStream) {
        List<ExcelDataResult> results = new ArrayList<>();

        // sheet文件流会自动关闭
        ExcelReader excelReader = EasyExcel.read(inputStream).build();

        final List<ExcelDataListener> listeners = getExcelDataListener(1000);

        final List<ReadSheet> readSheet = getReadSheet(listeners);
        // read multi sheet handle
        excelReader.read(readSheet);

        int i = 0;
        for (ExcelDataListener listener : listeners) {
            final List dataList = listener.getDataList();
            if (CollUtil.isEmpty(dataList)){
                continue;
            }
            String sheetName = readSheet.get(i).getSheetName();
            results.add(new ExcelDataResult(sheetName, listener.getHeadMap(), listener.getDataList()));
            i++;
        }
        return results;
    }

    private static List<ExcelDataListener> getExcelDataListener(int n) {
        List<ExcelDataListener> listeners = new ArrayList<>();
        for (int i = 0; i < n; i++) {
            final ExcelDataListener excelDataListener = new ExcelDataListener();
            listeners.add(excelDataListener);
        }
        return listeners;
    }

    public static List<ReadSheet> getReadSheet(List<ExcelDataListener> listeners){
        List<ReadSheet> sheets = new ArrayList<>();
        for (int i = 0; i < listeners.size(); i++) {
            try {
                ReadSheet sheet1 = EasyExcel.readSheet(i).headRowNumber(1).registerReadListener(listeners.get(i)).build();
                sheets.add(sheet1);
            } catch (Exception e) {

            }
        }

        return sheets;
    }


	// 如果传入类，会自动映射
	public static <T> List<T> simpleRead(InputStream inputStream, Class<T> clazz) {
        ExcelDataListener<T> dataListener = new ExcelDataListener<>();
        EasyExcel.read(inputStream, clazz, dataListener).sheet().doRead();
        return dataListener.getDataList();
    }


    public static void generatorExcel(String sheetName, Map<Integer, String> head, Map<Integer, String> head2, List<Map<Integer, String>> datas, String path){
        // row data
        System.out.println(head);

        List<Map<Integer, String>> newData = new ArrayList<>();

        final ExcelFillCellMergePrevColUtils mergePrevColUtils = new ExcelFillCellMergePrevColUtils();

        if (CollUtil.isNotEmpty(head2)){
            newData.add(head2);
        }

        if (CollUtil.isNotEmpty(head)){
            // first not null key
            Integer max = null;
            for (Integer key : head.keySet()) {
                if (key == 0){
                    continue;
                }
                final String s = head.get(key);
                if (StringUtils.isNotBlank(s)){
                    max = key - 1;
                    break;
                }
            }
            if (max == null){
                max = (int) head.keySet().stream().mapToLong(Integer::longValue).max().orElse(0);
            }
            mergePrevColUtils.add(0, 0, max);
        }

        if (CollUtil.isNotEmpty(datas)){
            newData.addAll(datas);
            // last row merge
            final Map<Integer, String> row = datas.get(datas.size() - 1);
            // first not null key
            Integer max = null;
            for (Integer key : row.keySet()) {
                if (key == 0){
                    continue;
                }
                final String s = row.get(key);
                if (StringUtils.isNotBlank(s)){
                    max = key - 1;
                    break;
                }
            }
            if (max == null){
                max = (int) row.keySet().stream().mapToLong(Integer::longValue).max().orElse(0);
            }
            mergePrevColUtils.add(datas.size() + 1, 0, max);
        }

        List<List<String>> headerList = new ArrayList<>();
        for (Integer integer : head.keySet()) {
            final String value = head.get(integer);
            if(StringUtils.isBlank(value) || "null".equals(value)){
                continue;
            }else {
                headerList.add(Collections.singletonList(head.get(integer)));
            }
        }

        // set style
        HorizontalCellStyleStrategy horizontalCellStyleStrategy = getHorizontalCellStyleStrategy();

        EasyExcel.write(new File(path)).sheet(StringUtils.isBlank(sheetName) ? "sheet1" : sheetName)
                .head(headerList)
                .registerWriteHandler(mergePrevColUtils)
                .registerWriteHandler(new ExcelRowHeightStyleStrategy(48, 32 , 32))
                .registerWriteHandler(new ExcelWidthStyleStrategy())
                .registerWriteHandler(horizontalCellStyleStrategy).doWrite(newData);
    }

    private static HorizontalCellStyleStrategy getHorizontalCellStyleStrategy() {
        WriteCellStyle headWriteCellStyle = new WriteCellStyle();
        //设置背景颜色
        headWriteCellStyle.setFillForegroundColor(IndexedColors.WHITE.getIndex());
        //设置头字体
        WriteFont headWriteFont = new WriteFont();
        headWriteFont.setFontHeightInPoints((short)16);
        headWriteFont.setBold(true);
        headWriteFont.setFontName("宋体");
        headWriteCellStyle.setWriteFont(headWriteFont);
        //设置头居中
        headWriteCellStyle.setHorizontalAlignment(HorizontalAlignment.CENTER);
        headWriteCellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        headWriteCellStyle.setBorderLeft(BorderStyle.THIN);
        headWriteCellStyle.setBorderRight(BorderStyle.THIN);
        headWriteCellStyle.setBorderTop(BorderStyle.THIN);
        headWriteCellStyle.setBorderBottom(BorderStyle.THIN);

        // 内容策略
        WriteCellStyle contentWriteCellStyle = new WriteCellStyle();
        //设置 水平居中
        contentWriteCellStyle.setHorizontalAlignment(HorizontalAlignment.CENTER);
        contentWriteCellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        WriteFont headWriteFont2 = new WriteFont();
        headWriteFont2.setFontHeightInPoints((short)11);
        headWriteFont2.setBold(false);
        headWriteFont2.setFontName("宋体");
        contentWriteCellStyle.setWriteFont(headWriteFont2);
        contentWriteCellStyle.setBorderLeft(BorderStyle.THIN);
        contentWriteCellStyle.setBorderRight(BorderStyle.THIN);
        contentWriteCellStyle.setBorderTop(BorderStyle.THIN);
        contentWriteCellStyle.setBorderBottom(BorderStyle.THIN);

        HorizontalCellStyleStrategy horizontalCellStyleStrategy = new HorizontalCellStyleStrategy(headWriteCellStyle, contentWriteCellStyle);
        return horizontalCellStyleStrategy;
    }


}

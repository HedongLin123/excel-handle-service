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

/**
 * @Description
 * @Author IT动力
 * @Date 2023/04/20 11:39
 */
@Slf4j
public class ExcelParse {

    /**
     * 读取excel，返回多个sheet的读取结果
     * @param inputStream
     * @return
     */
    public static List<ExcelDataResult> simpleRead(InputStream inputStream) {
        List<ExcelDataResult> results = new ArrayList<>();

        // sheet文件流会自动关闭
        ExcelReader excelReader = EasyExcel.read(inputStream).build();

        // 创建监听器1000，因为不知道sheet的个数，这里最多处理1000个sheet
        final List<ExcelDataListener> listeners = getExcelDataListener(1000);

        // 根据监听器，读取1000个sheet
        final List<ReadSheet> readSheet = getReadSheet(listeners);

        // 进行读取
        excelReader.read(readSheet);

        int i = 0;
        for (ExcelDataListener listener : listeners) {
            // 如果监听器监听的数据没有，表示这个sheet的内容为空，则不处理
            final List dataList = listener.getDataList();
            if (CollUtil.isEmpty(dataList)){
                continue;
            }
            // 注意：获取sheet名称，实际上获取结果为空，不知道为什么
            String sheetName = readSheet.get(i).getSheetName();
            // 将实际获取到的sheet结果数据放入results
            results.add(new ExcelDataResult(sheetName, listener.getHeadMap(), listener.getDataList()));
            i++;
        }
        return results;
    }

    /**
     * 创建n个监听器
     * @param n
     * @return
     */
    private static List<ExcelDataListener> getExcelDataListener(int n) {
        List<ExcelDataListener> listeners = new ArrayList<>();
        for (int i = 0; i < n; i++) {
            final ExcelDataListener excelDataListener = new ExcelDataListener();
            listeners.add(excelDataListener);
        }
        return listeners;
    }

    /**
     * 根据n个监听器，读取对应sheet，如果sheet时不存在的，读取结果一定是为空，实际只要判断sheet为空不处理即可动态处理不同的sheet
     * @param listeners
     * @return
     */
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


    /**
     * 生成excel数据
     * @param sheetName sheet名称
     * @param head 标题
     * @param head2 表头
     * @param datas 数据
     * @param path 生成excel路径
     */
    public static void generatorExcel(String sheetName, Map<Integer, String> head, Map<Integer, String> head2, List<Map<Integer, String>> datas, String path){
        // row data
//        System.out.println(head);

        List<Map<Integer, String>> newData = new ArrayList<>();

        final ExcelFillCellMergePrevColUtils mergePrevColUtils = new ExcelFillCellMergePrevColUtils();

        // 将第二行表头的数据假如数据列表，不做处理
        if (CollUtil.isNotEmpty(head2)){
            newData.add(head2);
        }

        // 标题需要进行行合并并居中
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

        // 内容的最后一行处理，因为最后一行是小计，从小计开始，单元格向右移动，第一个不为空的数据的前一个单元格与小计单元格合并
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

        // 标题数据处理
        List<List<String>> headerList = new ArrayList<>();
        for (Integer integer : head.keySet()) {
            final String value = head.get(integer);
            if(StringUtils.isBlank(value) || "null".equals(value)){
                continue;
            }else {
                headerList.add(Collections.singletonList(head.get(integer)));
            }
        }

        // 设置单元格样式
        HorizontalCellStyleStrategy horizontalCellStyleStrategy = getHorizontalCellStyleStrategy();

        // 生成excel
        EasyExcel.write(new File(path))
                // 设置sheet名称
                .sheet(StringUtils.isBlank(sheetName) ? "sheet1" : sheetName)
                // 设置标题
                .head(headerList)
                // 合并单元格处理
                .registerWriteHandler(mergePrevColUtils)
                // 设置行高处理
                .registerWriteHandler(new ExcelRowHeightStyleStrategy(48, 32 , 32))
                // 设置单元格宽度处理
                .registerWriteHandler(new ExcelWidthStyleStrategy())
                // 设置样式处理
                .registerWriteHandler(horizontalCellStyleStrategy)
                .doWrite(newData);
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
        // 设置边框
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
        // 设置边框
        contentWriteCellStyle.setBorderLeft(BorderStyle.THIN);
        contentWriteCellStyle.setBorderRight(BorderStyle.THIN);
        contentWriteCellStyle.setBorderTop(BorderStyle.THIN);
        contentWriteCellStyle.setBorderBottom(BorderStyle.THIN);

        return new HorizontalCellStyleStrategy(headWriteCellStyle, contentWriteCellStyle);
    }


}

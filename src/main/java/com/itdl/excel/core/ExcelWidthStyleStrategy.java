package com.itdl.excel.core;

import com.alibaba.excel.metadata.Head;
import com.alibaba.excel.metadata.data.WriteCellData;
import com.alibaba.excel.write.metadata.holder.WriteSheetHolder;
import com.alibaba.excel.write.style.column.AbstractColumnWidthStyleStrategy;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;

import java.util.List;

 /**
 * @Description 自定义设置列宽
 * @Author IT动力
 * @Date 2023/04/20 11:39
 */
public class ExcelWidthStyleStrategy extends AbstractColumnWidthStyleStrategy {

    // 统计setColumnWidth被调用多少次
    private static int count = 0;

    @Override
    protected void setColumnWidth(WriteSheetHolder writeSheetHolder, List<WriteCellData<?>> cellDataList, Cell cell, Head head, Integer relativeRowIndex, Boolean isHead) {
        // 简单设置，首先判断是否存在列，存在则设置特定列的宽度
        Sheet sheet = writeSheetHolder.getSheet();
        if (cell.getColumnIndex() == 1){
            sheet.setColumnWidth(cell.getColumnIndex(), 14 * 256);
        }
        if (cell.getColumnIndex() == 2){
            sheet.setColumnWidth(cell.getColumnIndex(), 18 * 256);
        }
        if (cell.getColumnIndex() == 3){
            sheet.setColumnWidth(cell.getColumnIndex(), 10 * 256);
        }
        if (cell.getColumnIndex() == 4){
            sheet.setColumnWidth(cell.getColumnIndex(), 10 * 256);
        }
        if (cell.getColumnIndex() == 5){
            sheet.setColumnWidth(cell.getColumnIndex(), 10 * 256);
        }
        if (cell.getColumnIndex() == 6){
            sheet.setColumnWidth(cell.getColumnIndex(), 10 * 256);
        }
        if (cell.getColumnIndex() == 7){
            sheet.setColumnWidth(cell.getColumnIndex(), 10 * 256);
        }
        if (cell.getColumnIndex() == 8){
            sheet.setColumnWidth(cell.getColumnIndex(), 10 * 256);
        }
        if (cell.getColumnIndex() == 9){
            sheet.setColumnWidth(cell.getColumnIndex(), 16 * 256);
        }
        if (cell.getColumnIndex() == 10){
            sheet.setColumnWidth(cell.getColumnIndex(), 10 * 256);
        }
        if (cell.getColumnIndex() == 11){
            sheet.setColumnWidth(cell.getColumnIndex(), 16 * 256);
        }
        if (cell.getColumnIndex() == 12){
            sheet.setColumnWidth(cell.getColumnIndex(), 10 * 256);
        }
        if (cell.getColumnIndex() == 13){
            sheet.setColumnWidth(cell.getColumnIndex(), 10 * 256);
        }
        if (cell.getColumnIndex() == 14){
            sheet.setColumnWidth(cell.getColumnIndex(), 10 * 256);
        }
        if (cell.getColumnIndex() == 15){
            sheet.setColumnWidth(cell.getColumnIndex(), 10 * 256);
        }
    }
}


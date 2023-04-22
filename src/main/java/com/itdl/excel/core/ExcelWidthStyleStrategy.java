package com.itdl.excel.core;

import com.alibaba.excel.metadata.Head;
import com.alibaba.excel.metadata.data.WriteCellData;
import com.alibaba.excel.write.metadata.holder.WriteSheetHolder;
import com.alibaba.excel.write.style.column.AbstractColumnWidthStyleStrategy;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;

import java.util.List;

/**
 * 自动设置列宽
 *
 * @date 01/22/2021 14:53
 */
public class ExcelWidthStyleStrategy extends AbstractColumnWidthStyleStrategy {

    // 统计setColumnWidth被调用多少次
    private static int count = 0;

    @Override
    protected void setColumnWidth(WriteSheetHolder writeSheetHolder, List<WriteCellData<?>> cellDataList, Cell cell, Head head, Integer relativeRowIndex, Boolean isHead) {
        // 简单设置
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


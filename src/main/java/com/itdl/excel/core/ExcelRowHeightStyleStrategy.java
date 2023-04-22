package com.itdl.excel.core;

import com.alibaba.excel.write.style.row.AbstractRowHeightStyleStrategy;
import org.apache.poi.ss.usermodel.Row;

/**
 * @Description
 * @Author donglin.he
 * @Date 2023/04/20 16:34
 */
public class ExcelRowHeightStyleStrategy extends AbstractRowHeightStyleStrategy {
    // title height
    double headHeight;
    // seond row 表头的行高
    double secondHeight;
    // 内容的行高
    double contentHeight;

    public ExcelRowHeightStyleStrategy(double headHeight, double secondHeight, double contentHeight) {
        this.headHeight = headHeight;
        this.secondHeight = secondHeight;
        this.contentHeight = contentHeight;
    }

    @Override
    protected void setHeadColumnHeight(Row row, int i) {
        row.setHeightInPoints((float) headHeight);
    }

    @Override
    protected void setContentColumnHeight(Row row, int i) {
        // second row set header style
        final int rowNum = row.getRowNum();
        if (rowNum == 1){
            row.setHeightInPoints((float) secondHeight);
        }else {
            row.setHeightInPoints((float) contentHeight);
        }
    }
}

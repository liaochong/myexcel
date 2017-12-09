package com.github.liaochong.html2excel.core.style;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * @author liaochong
 * @version 1.0
 */
public interface CellStyleFactory {

    /**
     * 单元格样式提供
     *
     * @param workbook workbook
     * @return CellStyle
     */
    CellStyle supply(Workbook workbook);
}

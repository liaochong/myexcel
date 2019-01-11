package com.github.liaochong.html2excel.core;

import org.apache.poi.ss.usermodel.Workbook;

/**
 * @author liaochong
 * @version 1.0
 */
public interface ExcelFactory {

    /**
     * 是否使用默认样式
     *
     * @return ExcelFactory
     */
    ExcelFactory useDefaultStyle();

    /**
     * 窗口冻结
     *
     * @param freezePanes 窗口冻结区域
     * @return ExcelFactory
     */
    ExcelFactory freezePanes(FreezePane... freezePanes);

    /**
     * 设置workbookType为SXSSFWorkbook的内存数据保有量
     *
     * @param rowAccessWindowSize 内存数据保有量
     * @return ExcelFactory
     */
    ExcelFactory rowAccessWindowSize(int rowAccessWindowSize);

    /**
     * 设置workbook类型
     *
     * @param workbookType 工作簿类型
     * @return ExcelFactory
     */
    ExcelFactory workbookType(WorkbookType workbookType);

    /**
     * 构建
     *
     * @return workbook
     */
    Workbook build();
}

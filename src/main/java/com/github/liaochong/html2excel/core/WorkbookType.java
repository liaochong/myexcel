package com.github.liaochong.html2excel.core;

/**
 * @author liaochong
 * @version 1.0
 */
public enum WorkbookType {
    /**
     * .xls
     */
    XLS,
    /**
     * .xlsx
     */
    XLSX;

    public static boolean isXls(WorkbookType workbookType) {
        return XLS.equals(workbookType);
    }

}

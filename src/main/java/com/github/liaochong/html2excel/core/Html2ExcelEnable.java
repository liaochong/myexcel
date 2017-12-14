package com.github.liaochong.html2excel.core;

import java.util.Map;

import org.apache.poi.ss.usermodel.Workbook;

/**
 * @author liaochong
 * @version 1.0
 */
public interface Html2ExcelEnable {

    /**
     * excel类型
     * 
     * @param workbookType workbookType
     * @return Html2ExcelEnable
     */
    Html2ExcelEnable type(WorkbookType workbookType);

    /**
     * 使用默认样式
     * 
     * @return Html2ExcelEnable
     */
    Html2ExcelEnable useDefaultStyle();

    /**
     * 获取模板
     * 
     * @param path 模板路径
     * @return Html2ExcelEnable
     */
    Html2ExcelEnable getTemplate(String path);

    /**
     * 构建
     * 
     * @param renderData 渲染数据
     * @return Workbook
     */
    Workbook build(Map<String, Object> renderData);
}

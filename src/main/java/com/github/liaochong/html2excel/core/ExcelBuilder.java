package com.github.liaochong.html2excel.core;

import java.util.Map;

import org.apache.poi.ss.usermodel.Workbook;

/**
 * excel创建者接口
 * 
 * @author liaochong
 * @version 1.0
 */
public interface ExcelBuilder {

    /**
     * excel类型
     *
     * @param workbookType workbookType
     * @return ExcelBuilder
     */
    ExcelBuilder type(WorkbookType workbookType);

    /**
     * 使用默认样式
     *
     * @return ExcelBuilder
     */
    ExcelBuilder useDefaultStyle();

    /**
     * 获取模板
     *
     * @param path 模板路径
     * @return ExcelBuilder
     */
    ExcelBuilder getTemplate(String path);

    /**
     * 构建
     *
     * @param renderData 渲染数据
     * @return Workbook
     */
    Workbook build(Map<String, Object> renderData);
}

package com.github.liaochong.html2excel.core;

import java.io.File;
import java.io.IOException;
import java.util.Map;
import java.util.Objects;
import java.util.UUID;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Workbook;

import com.github.liaochong.html2excel.exception.ExcelBuildException;

/**
 * excel创建者接口
 * 
 * @author liaochong
 * @version 1.0
 */
public abstract class ExcelBuilder {

    protected HtmlToExcelFactory htmlToExcelFactory = new HtmlToExcelFactory();

    /**
     * excel类型
     *
     * @param workbookType workbookType
     * @return ExcelBuilder
     */
    public ExcelBuilder workbookType(WorkbookType workbookType) {
        htmlToExcelFactory.workbookType(workbookType);
        return this;
    }

    /**
     * 使用默认样式
     *
     * @return ExcelBuilder
     */
    public ExcelBuilder useDefaultStyle() {
        htmlToExcelFactory.useDefaultStyle();
        return this;
    }

    /**
     * 获取模板
     *
     * @param path 模板路径
     * @return ExcelBuilder
     */
    public abstract ExcelBuilder getTemplate(String path);

    /**
     * 构建
     *
     * @param renderData 渲染数据
     * @return Workbook
     */
    public abstract Workbook build(Map<String, Object> renderData);

    /**
     * 分离文件路径
     * 
     * @param path 文件路径
     * @return String[]
     */
    String[] splitFilePath(String path) {
        if (StringUtils.isBlank(path)) {
            throw new NullPointerException();
        }
        int lastPackageIndex = path.lastIndexOf("/");
        if (lastPackageIndex == -1 || lastPackageIndex == path.length() - 1) {
            throw new IllegalArgumentException();
        }
        String basePackagePath = path.substring(0, lastPackageIndex);
        String templateName = path.substring(lastPackageIndex);
        return new String[] { basePackagePath, templateName };
    }

    /**
     * 依据前缀名称创建临时文件
     * 
     * @param prefix 临时文件前缀
     * @return File
     */
    File createTempFile(String prefix) {
        try {
            return File.createTempFile(prefix + UUID.randomUUID(), ".html");
        } catch (IOException e) {
            throw ExcelBuildException.of("failed to create temp html file", e);
        }
    }

    /**
     * 删除临时文件
     * 
     * @param file 临时文件
     */
    void deleteTempFile(File file) {
        if (Objects.isNull(file) || !file.exists()) {
            return;
        }
        boolean isDeleted = file.delete();
        if (!isDeleted) {
            throw new IllegalStateException("failed to delete temp html file");
        }
    }

}

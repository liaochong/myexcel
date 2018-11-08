/*
 * Copyright 2017 the original author or authors.
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *      http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
package com.github.liaochong.html2excel.core;

import com.github.liaochong.html2excel.exception.ExcelBuildException;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.File;
import java.io.IOException;
import java.util.Map;
import java.util.Objects;
import java.util.UUID;

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
     * 选择固定区域
     *
     * @param freezePanes 固定区域
     * @return ExcelBuilder
     */
    public ExcelBuilder freezePanes(FreezePane... freezePanes) {
        if (Objects.isNull(freezePanes) || freezePanes.length == 0) {
            return this;
        }
        htmlToExcelFactory.freezePanes(freezePanes);
        return this;
    }

    /**
     * 设置模板
     *
     * @param path 模板路径
     * @return ExcelBuilder
     */
    public abstract ExcelBuilder template(String path);

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
        if (Objects.isNull(path)) {
            throw new NullPointerException();
        }
        int lastPackageIndex = path.lastIndexOf("/");
        if (lastPackageIndex == -1 || lastPackageIndex == path.length() - 1) {
            throw new IllegalArgumentException();
        }
        String basePackagePath = path.substring(0, lastPackageIndex);
        String templateName = path.substring(lastPackageIndex);
        return new String[]{basePackagePath, templateName};
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

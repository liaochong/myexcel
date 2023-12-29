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
package com.github.liaochong.myexcel.core;

import com.github.liaochong.myexcel.core.strategy.AutoWidthStrategy;
import com.github.liaochong.myexcel.core.strategy.SheetStrategy;
import com.github.liaochong.myexcel.core.strategy.WidthStrategy;
import com.github.liaochong.myexcel.core.templatehandler.TemplateHandler;
import com.github.liaochong.myexcel.exception.ExcelBuildException;
import com.github.liaochong.myexcel.utils.ReflectUtil;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.IOException;
import java.util.List;
import java.util.Map;

/**
 * excel创建者接口
 *
 * @author liaochong
 * @version 1.0
 */
public abstract class AbstractExcelBuilder implements ExcelBuilder {

    protected TemplateHandler templateHandler;

    protected HtmlToExcelFactory htmlToExcelFactory = new HtmlToExcelFactory();

    protected AbstractExcelBuilder(Class<? extends TemplateHandler> templateHandlerClass) {
        widthStrategy(WidthStrategy.COMPUTE_AUTO_WIDTH);
        sheetStrategy(SheetStrategy.MULTI_SHEET);
        this.templateHandler = ReflectUtil.newInstance(templateHandlerClass);
    }

    @Override
    public AbstractExcelBuilder workbookType(WorkbookType workbookType) {
        htmlToExcelFactory.workbookType(workbookType);
        return this;
    }

    @Override
    public AbstractExcelBuilder useDefaultStyle() {
        htmlToExcelFactory.useDefaultStyle();
        return this;
    }

    @Override
    public ExcelBuilder applyDefaultStyle() {
        htmlToExcelFactory.applyDefaultStyle();
        return this;
    }

    @Override
    public AbstractExcelBuilder widthStrategy(WidthStrategy widthStrategy) {
        htmlToExcelFactory.widthStrategy(widthStrategy);
        return this;
    }

    @Deprecated
    @Override
    public AbstractExcelBuilder autoWidthStrategy(AutoWidthStrategy autoWidthStrategy) {
        htmlToExcelFactory.widthStrategy(AutoWidthStrategy.map(autoWidthStrategy));
        return this;
    }

    @Override
    public AbstractExcelBuilder sheetStrategy(SheetStrategy sheetStrategy) {
        htmlToExcelFactory.sheetStrategy(sheetStrategy);
        return this;
    }

    @Override
    public AbstractExcelBuilder freezePanes(FreezePane... freezePanes) {
        if (freezePanes == null || freezePanes.length == 0) {
            return this;
        }
        htmlToExcelFactory.freezePanes(freezePanes);
        return this;
    }

    @Override
    public ExcelBuilder nameManager(Map<String, List<?>> nameMapping) {
        htmlToExcelFactory.nameManager(nameMapping);
        return this;
    }

    /**
     * 构建
     *
     * @param data 模板参数
     * @return Workbook
     */
    @Override
    public <T> Workbook build(Map<String, T> data) {
        String template = templateHandler.render(data);
        try {
            return HtmlToExcelFactory.readHtml(template, htmlToExcelFactory).build();
        } catch (Exception e) {
            throw new ExcelBuildException("Build excel failure", e);
        }
    }

    @Override
    public ExcelBuilder classpathTemplate(String path) {
        templateHandler.classpathTemplate(path);
        return this;
    }

    @Deprecated
    @Override
    public ExcelBuilder template(String path) {
        return classpathTemplate(path);
    }

    @Override
    public ExcelBuilder fileTemplate(String dirPath, String fileName) {
        templateHandler.fileTemplate(dirPath, fileName);
        return this;
    }

    @Override
    public void close() throws IOException {
        if (htmlToExcelFactory != null) {
            htmlToExcelFactory.closeWorkbook();
        }
    }
}

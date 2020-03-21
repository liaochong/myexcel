/*
 * Copyright 2019 liaochong
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 * http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
package com.github.liaochong.myexcel.core;

import com.github.liaochong.myexcel.core.strategy.WidthStrategy;
import com.github.liaochong.myexcel.core.templatehandler.TemplateHandler;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.Closeable;
import java.io.IOException;
import java.util.Map;

/**
 * @author liaochong
 * @version 1.0
 */
public class TemplateExcelBuilder implements Closeable {

    private ExcelBuilder excelBuilder;

    private TemplateExcelBuilder(ExcelBuilder excelBuilder) {
        this.excelBuilder = excelBuilder;
    }

    public static TemplateExcelBuilder of(Class<? extends TemplateHandler> templateHandlerClass) {
        return new TemplateExcelBuilder(new DefaultTemplateExcelBuilder(templateHandlerClass));
    }

    public TemplateExcelBuilder workbookType(WorkbookType workbookType) {
        excelBuilder.workbookType(workbookType);
        return this;
    }

    public TemplateExcelBuilder useDefaultStyle() {
        excelBuilder.useDefaultStyle();
        return this;
    }

    public TemplateExcelBuilder widthStrategy(WidthStrategy widthStrategy) {
        excelBuilder.widthStrategy(widthStrategy);
        return this;
    }

    public TemplateExcelBuilder freezePanes(FreezePane... freezePanes) {
        return this;
    }

    public TemplateExcelBuilder classpathTemplate(String path) {
        excelBuilder.classpathTemplate(path);
        return this;
    }

    public TemplateExcelBuilder fileTemplate(String dirPath, String fileName) {
        excelBuilder.fileTemplate(dirPath, fileName);
        return this;
    }

    public <E> Workbook build(Map<String, E> data) {
        return excelBuilder.build(data);
    }

    @Override
    public void close() throws IOException {
        excelBuilder.close();
    }
}

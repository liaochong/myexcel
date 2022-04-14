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

import com.github.liaochong.myexcel.core.templatehandler.TemplateHandler;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.Closeable;
import java.io.IOException;
import java.util.Map;

/**
 * @author liaochong
 * @version 4.2.0
 */
public class TemplateStreamExcelBuilder implements Closeable {

    private DefaultStreamExcelBuilder<Void> defaultStreamExcelBuilder;

    private TemplateStreamExcelBuilder() {
    }

    public static TemplateStreamExcelBuilder of(Class<? extends TemplateHandler> templateHandlerClass) {
        TemplateStreamExcelBuilder templateStreamExcelBuilder = new TemplateStreamExcelBuilder();
        templateStreamExcelBuilder.defaultStreamExcelBuilder = DefaultStreamExcelBuilder.of(Void.class)
                .templateHandler(templateHandlerClass)
                .start();
        return templateStreamExcelBuilder;
    }

    public void append(String templateFilePath, Map<String, Object> renderData) {
        defaultStreamExcelBuilder.append(templateFilePath, renderData);
    }

    public <E> void append(String templateDir, String templateFileName, Map<String, E> renderData) {
        defaultStreamExcelBuilder.append(templateDir, templateFileName, renderData);
    }

    public Workbook build() {
        return defaultStreamExcelBuilder.build();
    }

    @Override
    public void close() throws IOException {
        defaultStreamExcelBuilder.close();
    }
}

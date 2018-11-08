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
import groovy.lang.Writable;
import groovy.text.Template;
import groovy.text.markup.MarkupTemplateEngine;
import groovy.text.markup.TemplateConfiguration;
import org.apache.commons.codec.CharEncoding;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.io.OutputStreamWriter;
import java.io.Reader;
import java.io.Writer;
import java.util.Map;

/**
 * @author liaochong
 * @version 1.0
 */
public class GroovyExcelBuilder extends ExcelBuilder {

    private Template template;

    @Override
    public ExcelBuilder setTemplate(String path) {
        TemplateConfiguration config = new TemplateConfiguration();
        MarkupTemplateEngine engine = new MarkupTemplateEngine(config);
        try {
            InputStream is = Thread.currentThread().getContextClassLoader().getResourceAsStream(path);
            Reader reader = new InputStreamReader(is);
            template = engine.createTemplate(reader);
            return this;
        } catch (ClassNotFoundException | IOException e) {
            throw ExcelBuildException.of("failed to get groovy template", e);
        }
    }

    @Override
    public Workbook build(Map<String, Object> renderData) {
        try {
            File htmlFile = this.createTempFile("groovy_temp_");
            Writer out = new BufferedWriter(new OutputStreamWriter(new FileOutputStream(htmlFile), CharEncoding.UTF_8));

            Writable output = template.make(renderData);
            output.writeTo(out);

            Workbook workbook = HtmlToExcelFactory.readHtml(htmlFile, htmlToExcelFactory).build();
            this.deleteTempFile(htmlFile);
            return workbook;
        } catch (Exception e) {
            throw ExcelBuildException.of("failed to build excel", e);
        }
    }
}

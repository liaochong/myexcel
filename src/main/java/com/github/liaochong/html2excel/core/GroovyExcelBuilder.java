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

import com.github.liaochong.html2excel.core.io.TempFileOperator;
import com.github.liaochong.html2excel.exception.ExcelBuildException;
import groovy.lang.Writable;
import groovy.text.Template;
import groovy.text.markup.MarkupTemplateEngine;
import groovy.text.markup.TemplateConfiguration;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.io.Reader;
import java.io.Writer;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Map;
import java.util.Objects;

/**
 * @author liaochong
 * @version 1.0
 */
public class GroovyExcelBuilder extends AbstractExcelBuilder {

    private Template template;

    @Override
    public ExcelBuilder template(String path) {
        TemplateConfiguration config = new TemplateConfiguration();
        MarkupTemplateEngine engine = new MarkupTemplateEngine(config);
        try (InputStream is = Thread.currentThread().getContextClassLoader().getResourceAsStream(path);
             Reader reader = new InputStreamReader(is, StandardCharsets.UTF_8)) {
            template = engine.createTemplate(reader);
            return this;
        } catch (ClassNotFoundException | IOException e) {
            throw ExcelBuildException.of("Failed to get groovy template", e);
        }
    }

    @Override
    public <T> Workbook build(Map<String, T> renderData) {
        Objects.requireNonNull(template, "The template cannot be empty. Please set the template first.");
        Path htmlFile = tempFileOperator.createTempFile("groovy_temp_", TempFileOperator.HTML_SUFFIX);
        try (Writer out = Files.newBufferedWriter(htmlFile, StandardCharsets.UTF_8)) {
            Writable output = template.make(renderData);
            output.writeTo(out);
            return HtmlToExcelFactory.readHtml(htmlFile.toFile(), htmlToExcelFactory).build();
        } catch (Exception e) {
            throw ExcelBuildException.of("Failed to build excel", e);
        } finally {
            tempFileOperator.deleteTempFile();
        }
    }
}

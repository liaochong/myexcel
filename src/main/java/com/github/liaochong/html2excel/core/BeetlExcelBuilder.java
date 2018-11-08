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
import org.apache.commons.codec.CharEncoding;
import org.apache.poi.ss.usermodel.Workbook;
import org.beetl.core.Configuration;
import org.beetl.core.GroupTemplate;
import org.beetl.core.Template;
import org.beetl.core.resource.ClasspathResourceLoader;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStreamWriter;
import java.io.Writer;
import java.nio.charset.StandardCharsets;
import java.util.Map;
import java.util.Objects;

/**
 * beetl excel创建者
 *
 * @author liaochong
 * @version 1.0
 */
public class BeetlExcelBuilder extends ExcelBuilder {

    private Template template;

    @Override
    public ExcelBuilder setTemplate(String path) {
        try {
            String[] filePath = this.splitFilePath(path);
            ClasspathResourceLoader resourceLoader = new ClasspathResourceLoader(filePath[0]);
            Configuration cfg = Configuration.defaultConfiguration();
            cfg.setCharset(CharEncoding.UTF_8);
            GroupTemplate gt = new GroupTemplate(resourceLoader, cfg);
            template = gt.getTemplate(filePath[1]);
            return this;
        } catch (IOException e) {
            throw ExcelBuildException.of("Failed to get beetl template", e);
        }
    }

    @Override
    public Workbook build(Map<String, Object> renderData) {
        Objects.requireNonNull(template, "The template cannot be empty. Please set the template first.");
        try {
            File htmlFile = this.createTempFile("beetl_temp_");
            Writer out = new BufferedWriter(new OutputStreamWriter(new FileOutputStream(htmlFile), StandardCharsets.UTF_8));

            template.binding(renderData);
            template.renderTo(out);
            Workbook workbook = HtmlToExcelFactory.readHtml(htmlFile, htmlToExcelFactory).build();
            this.deleteTempFile(htmlFile);
            return workbook;
        } catch (Exception e) {
            throw ExcelBuildException.of("Failed to build excel", e);
        }
    }
}

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

import com.github.liaochong.myexcel.core.io.TempFileOperator;
import com.github.liaochong.myexcel.core.strategy.AutoWidthStrategy;
import com.github.liaochong.myexcel.exception.ExcelBuildException;
import org.apache.commons.codec.CharEncoding;
import org.apache.poi.ss.usermodel.Workbook;
import org.beetl.core.Configuration;
import org.beetl.core.GroupTemplate;
import org.beetl.core.Template;
import org.beetl.core.resource.ClasspathResourceLoader;

import java.io.IOException;
import java.io.Writer;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Map;
import java.util.Objects;

/**
 * beetl excel创建者
 *
 * @author liaochong
 * @version 1.0
 */
public class BeetlExcelBuilder extends AbstractExcelBuilder {

    private static final GroupTemplate GROUP_TEMPLATE;

    static {
        ClasspathResourceLoader resourceLoader = new ClasspathResourceLoader();
        Configuration cfg;
        try {
            cfg = Configuration.defaultConfiguration();
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
        cfg.setCharset(CharEncoding.UTF_8);
        GROUP_TEMPLATE = new GroupTemplate(resourceLoader, cfg);
    }

    private Template template;

    public BeetlExcelBuilder() {
        autoWidthStrategy(AutoWidthStrategy.AUTO_WIDTH);
    }

    @Override
    public ExcelBuilder template(String path) {
        template = GROUP_TEMPLATE.getTemplate(path);
        return this;
    }

    @Override
    public <T> Workbook build(Map<String, T> renderData) {
        Objects.requireNonNull(template, "The template cannot be empty. Please set the template first.");
        Path htmlFile = tempFileOperator.createTempFile("beetl_temp_", TempFileOperator.HTML_SUFFIX);
        try (Writer out = Files.newBufferedWriter(htmlFile, StandardCharsets.UTF_8)) {
            template.binding(renderData);
            template.renderTo(out);
            return HtmlToExcelFactory.readHtml(htmlFile.toFile(), htmlToExcelFactory).build();
        } catch (Exception e) {
            throw ExcelBuildException.of("Failed to build excel", e);
        } finally {
            tempFileOperator.deleteTempFile();
        }
    }
}

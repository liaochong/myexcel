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
import com.github.liaochong.myexcel.exception.ExcelBuildException;
import freemarker.template.Configuration;
import freemarker.template.Template;
import freemarker.template.TemplateExceptionHandler;
import org.apache.commons.codec.CharEncoding;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.IOException;
import java.io.Writer;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Map;
import java.util.Objects;

/**
 * freemarker的excel创建者
 *
 * @author liaochong
 * @version 1.0
 */
public class FreemarkerExcelBuilder extends AbstractExcelBuilder {

    private Template template;

    /**
     * 设置模板信息
     *
     * @param path 模板路径，相对路径
     */
    @Override
    public ExcelBuilder template(String path) {
        try {
            Configuration cfg = new Configuration(Configuration.VERSION_2_3_23);
            cfg.setTemplateExceptionHandler(TemplateExceptionHandler.RETHROW_HANDLER);
            cfg.setDefaultEncoding(CharEncoding.UTF_8);

            String[] filePath = this.splitFilePath(path);
            cfg.setClassLoaderForTemplateLoading(Thread.currentThread().getContextClassLoader(), filePath[0]);
            template = cfg.getTemplate(filePath[1]);
            return this;
        } catch (IOException e) {
            throw ExcelBuildException.of("Failed to get freemarker template", e);
        }
    }

    /**
     * 构建
     *
     * @param data 模板参数
     * @return Workbook
     */
    @Override
    public <T> Workbook build(Map<String, T> data) {
        Objects.requireNonNull(template, "The template cannot be empty. Please set the template first.");
        Path htmlFile = tempFileOperator.createTempFile("freemarker_temp_", TempFileOperator.HTML_SUFFIX);
        try (Writer out = Files.newBufferedWriter(htmlFile, StandardCharsets.UTF_8)) {
            template.process(data, out);
            return HtmlToExcelFactory.readHtml(htmlFile.toFile(), htmlToExcelFactory).build();
        } catch (Exception e) {
            throw ExcelBuildException.of("Failed to build excel", e);
        } finally {
            tempFileOperator.deleteTempFile();
        }
    }

}

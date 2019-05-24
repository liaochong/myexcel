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
import com.github.liaochong.myexcel.exception.ExcelBuildException;
import org.apache.commons.codec.CharEncoding;
import org.beetl.core.Configuration;
import org.beetl.core.GroupTemplate;
import org.beetl.core.Template;
import org.beetl.core.resource.ClasspathResourceLoader;

import java.io.IOException;
import java.io.Writer;
import java.util.Map;
import java.util.Objects;

/**
 * beetl excel创建者
 *
 * @author liaochong
 * @version 1.0
 */
public class BeetlExcelBuilder extends AbstractExcelBuilder {

    private Template template;

    public BeetlExcelBuilder() {
        autoWidthStrategy(AutoWidthStrategy.AUTO_WIDTH);
    }

    @Override
    public ExcelBuilder template(String path) {
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
    protected <T> void render(Map<String, T> data, Writer out) throws Exception {
        Objects.requireNonNull(template, "The template cannot be empty. Please set the template first.");
        template.binding(data);
        template.renderTo(out);
    }
}

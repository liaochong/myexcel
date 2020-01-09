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

import com.github.liaochong.myexcel.core.strategy.WidthStrategy;
import org.apache.commons.codec.CharEncoding;
import org.beetl.core.Configuration;
import org.beetl.core.GroupTemplate;
import org.beetl.core.Template;
import org.beetl.core.resource.ClasspathResourceLoader;

import java.io.IOException;
import java.io.Writer;
import java.util.Map;

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
        widthStrategy(WidthStrategy.AUTO_WIDTH);
    }

    @Override
    public ExcelBuilder template(String path) {
        template = GROUP_TEMPLATE.getTemplate(path);
        return this;
    }

    @Override
    protected <T> void render(Map<String, T> data, Writer out) {
        checkTemplate(template);
        template.binding(data);
        template.renderTo(out);
    }
}

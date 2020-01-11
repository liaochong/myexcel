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

import com.github.liaochong.myexcel.core.constant.Constants;
import lombok.extern.slf4j.Slf4j;
import org.thymeleaf.TemplateEngine;
import org.thymeleaf.context.Context;
import org.thymeleaf.templateresolver.ClassLoaderTemplateResolver;

import java.io.Writer;
import java.nio.charset.StandardCharsets;
import java.util.Map;
import java.util.Objects;

/**
 * ThymeleafExcelBuilder
 *
 * @author liaochong
 * @version 1.0
 */
@Slf4j
public class ThymeleafExcelBuilder extends AbstractExcelBuilder {

    private static final TemplateEngine TEMPLATE_ENGINE;

    private String filePath;

    static {
        TEMPLATE_ENGINE = new TemplateEngine();
        ClassLoaderTemplateResolver classLoaderTemplateResolver = new ClassLoaderTemplateResolver();
        classLoaderTemplateResolver.setCharacterEncoding(StandardCharsets.UTF_8.name());
        classLoaderTemplateResolver.setCacheable(true);
        TEMPLATE_ENGINE.setTemplateResolver(classLoaderTemplateResolver);
    }

    @Override
    public ExcelBuilder template(String path) {
        Objects.requireNonNull(path);
        if (!path.endsWith(Constants.HTML_SUFFIX)) {
            throw new IllegalArgumentException("ThymeleafExcelBuilder only supports files suffixed with .html");
        }
        if (path.startsWith("/")) {
            path = path.substring(1);
        }
        filePath = path;
        return this;
    }

    @Override
    protected <T> void render(Map<String, T> renderData, Writer out) throws Exception {
        Context context = new Context();
        context.setVariables(renderData);
        TEMPLATE_ENGINE.process(filePath, context, out);
    }

}

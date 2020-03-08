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
import org.thymeleaf.templateresolver.FileTemplateResolver;
import org.thymeleaf.templateresolver.TemplateResolver;

import java.io.File;
import java.io.Writer;
import java.nio.charset.StandardCharsets;
import java.util.HashMap;
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

    private static final Map<String, TemplateEngine> CFG_MAP = new HashMap<>();

    private TemplateEngine templateEngine;

    private String filePath;

    @Override
    public ExcelBuilder classpathTemplate(String path) {
        if (!path.endsWith(Constants.HTML_SUFFIX)) {
            throw new IllegalArgumentException("ThymeleafExcelBuilder only supports files suffixed with .html");
        }
        if (path.startsWith("/")) {
            path = path.substring(1);
        }
        filePath = path;
        doSetEngine(CLASSPATH);
        return this;
    }

    @Override
    public ExcelBuilder fileTemplate(String dirPath, String fileName) {
        if (!fileName.endsWith(Constants.HTML_SUFFIX)) {
            throw new IllegalArgumentException("ThymeleafExcelBuilder only supports files suffixed with .html");
        }
        if (!dirPath.endsWith("/")) {
            dirPath = dirPath + File.separator;
        }
        filePath = fileName;
        doSetEngine(dirPath);
        return this;
    }

    @Override
    protected <T> void render(Map<String, T> renderData, Writer out) throws Exception {
        Context context = new Context();
        context.setVariables(renderData);
        templateEngine.process(filePath, context, out);
    }

    private void doSetEngine(String dirPath) {
        templateEngine = CFG_MAP.get(dirPath);
        if (templateEngine == null) {
            templateEngine = doGetEngine(dirPath);
        }
    }

    private synchronized TemplateEngine doGetEngine(String dirPath) {
        TemplateEngine templateEngine = CFG_MAP.get(dirPath);
        if (templateEngine != null) {
            return templateEngine;
        }
        templateEngine = new TemplateEngine();
        TemplateResolver templateResolver;
        if (Objects.equals(dirPath, CLASSPATH)) {
            templateResolver = new ClassLoaderTemplateResolver();
            templateResolver.setCacheable(true);
        } else {
            templateResolver = new FileTemplateResolver();
            templateResolver.setPrefix(dirPath);
            templateResolver.setCacheable(false);
        }
        templateResolver.setCharacterEncoding(StandardCharsets.UTF_8.name());
        templateEngine.setTemplateResolver(templateResolver);
        CFG_MAP.put(dirPath, templateEngine);
        return templateEngine;
    }
}

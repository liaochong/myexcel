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

import com.github.liaochong.myexcel.core.io.TempFileOperator;
import com.github.liaochong.myexcel.core.strategy.AutoWidthStrategy;
import com.github.liaochong.myexcel.exception.ExcelBuildException;
import lombok.extern.slf4j.Slf4j;
import org.thymeleaf.TemplateEngine;
import org.thymeleaf.context.Context;
import org.thymeleaf.templateresolver.FileTemplateResolver;

import java.io.Writer;
import java.net.URISyntaxException;
import java.net.URL;
import java.nio.charset.StandardCharsets;
import java.nio.file.Paths;
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

    private String realPath;

    static {
        TEMPLATE_ENGINE = new TemplateEngine();
        FileTemplateResolver fileTemplateResolver = new FileTemplateResolver();
        fileTemplateResolver.setCharacterEncoding(StandardCharsets.UTF_8.name());
        fileTemplateResolver.setTemplateMode("HTML5");
        fileTemplateResolver.setCacheable(true);
        TEMPLATE_ENGINE.setTemplateResolver(fileTemplateResolver);
    }

    public ThymeleafExcelBuilder() {
        autoWidthStrategy(AutoWidthStrategy.AUTO_WIDTH);
    }

    @Override
    public ExcelBuilder template(String path) {
        Objects.requireNonNull(path);
        if (!path.endsWith(TempFileOperator.HTML_SUFFIX)) {
            throw new IllegalArgumentException("ThymeleafExcelBuilder only supports files suffixed with .html");
        }
        if (path.startsWith("/")) {
            path = path.substring(1);
        }
        String filePath = path;
        try {
            URL url = Thread.currentThread().getContextClassLoader().getResource(filePath);
            if (Objects.isNull(url)) {
                throw new IllegalAccessException("File path " + filePath + " is not accessible");
            }
            realPath = Paths.get(url.toURI()).toAbsolutePath().toString();
            log.info("Template file:" + realPath);
        } catch (IllegalAccessException | URISyntaxException e) {
            throw ExcelBuildException.of("Failed to build ExcelBuilder", e);
        }

        return this;
    }

    @Override
    protected <T> void render(Map<String, T> renderData, Writer out) throws Exception {
        Context context = new Context();
        context.setVariables(renderData);
        TEMPLATE_ENGINE.process(realPath, context, out);
    }

}

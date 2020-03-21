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
package com.github.liaochong.myexcel.core.templatehandler;

import com.github.liaochong.myexcel.exception.ExcelBuildException;
import groovy.lang.Writable;
import groovy.text.Template;
import groovy.text.markup.MarkupTemplateEngine;
import groovy.text.markup.TemplateConfiguration;

import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.io.Reader;
import java.io.Writer;
import java.nio.charset.StandardCharsets;
import java.util.Map;
import java.util.function.Supplier;

/**
 * @author liaochong
 * @version 1.0
 */
public class GroovyTemplateHandler extends AbstractTemplateHandler<Template, Template> {

    private static final MarkupTemplateEngine ENGINE;

    static {
        TemplateConfiguration config = new TemplateConfiguration();
        ENGINE = new MarkupTemplateEngine(config);
    }

    @Override
    public GroovyTemplateHandler classpathTemplate(String path) {
        if (path.startsWith("/")) {
            path = path.substring(1);
        }
        try (InputStream is = Thread.currentThread().getContextClassLoader().getResourceAsStream(path);
             Reader reader = new InputStreamReader(is, StandardCharsets.UTF_8)) {
            templateEngine = ENGINE.createTemplate(reader);
            return this;
        } catch (ClassNotFoundException | IOException e) {
            throw ExcelBuildException.of("Failed to get groovy template", e);
        }
    }

    @Override
    public GroovyTemplateHandler fileTemplate(String dirPath, String fileName) {
        throw new UnsupportedOperationException();
    }

    @Override
    protected void setTemplateEngine(String dirPath, Supplier<Template> supplier, String fileName) {

    }

    @Override
    protected Template getConfiguration(String dirPath) {
        return null;
    }

    @Override
    protected <F> void render(Map<String, F> renderData, Writer out) throws Exception {
        Writable output = templateEngine.make(renderData);
        output.writeTo(out);
    }
}

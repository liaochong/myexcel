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

import com.github.liaochong.myexcel.core.parser.HtmlTableParser;
import com.github.liaochong.myexcel.core.parser.ParseConfig;
import com.github.liaochong.myexcel.core.parser.Table;
import com.github.liaochong.myexcel.exception.ExcelBuildException;

import java.io.StringWriter;
import java.io.Writer;
import java.util.List;
import java.util.Map;
import java.util.function.Supplier;

/**
 * @author liaochong
 * @version 1.0
 */
public abstract class AbstractTemplateHandler<T, E> implements TemplateHandler {

    protected E templateEngine;

    @Override
    public AbstractTemplateHandler<T, E> classpathTemplate(String path) {
        setTemplateEngine(CLASSPATH, () -> this.getConfiguration(CLASSPATH), path);
        return this;
    }

    @Override
    public AbstractTemplateHandler<T, E> fileTemplate(String dirPath, String fileName) {
        setTemplateEngine(dirPath, () -> this.getConfiguration(dirPath), fileName);
        return this;
    }

    protected abstract void setTemplateEngine(String dirPath, Supplier<T> supplier, String fileName);

    protected abstract T getConfiguration(String dirPath);

    @Override
    public <F> String render(Map<String, F> renderData) {
        String template;
        try (Writer out = new StringWriter()) {
            render(renderData, out);
            template = out.toString();
        } catch (Exception e) {
            throw ExcelBuildException.of("Failed to build excel", e);
        }
        return template;
    }

    protected abstract <F> void render(Map<String, F> renderData, Writer out) throws Exception;

    /**
     * 模板引擎渲染，返回渲染后的数据
     *
     * @param renderData 渲染数据
     * @param <F>        被渲染数据类型
     * @return 输出流，who create who close;
     */
    public <F> List<Table> render(Map<String, F> renderData, ParseConfig parseConfig) throws Exception {
        String template = render(renderData);
        return HtmlTableParser.of(template).getAllTable(parseConfig);
    }
}

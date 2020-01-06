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
import org.apache.velocity.Template;
import org.apache.velocity.VelocityContext;
import org.apache.velocity.app.Velocity;
import org.apache.velocity.runtime.resource.loader.ClasspathResourceLoader;

import java.io.Writer;
import java.util.Map;
import java.util.Objects;

/**
 * velocity的excel创建者
 *
 * @author gaokai
 * @version 1.0
 */
public class VelocityExcelBuilder extends AbstractExcelBuilder {

    static {
        Velocity.setProperty(Velocity.RESOURCE_LOADER, "classpath");
        Velocity.setProperty("classpath.resource.loader.class", ClasspathResourceLoader.class.getName());
        Velocity.setProperty(Velocity.ENCODING_DEFAULT, CharEncoding.UTF_8);
        Velocity.setProperty(Velocity.INPUT_ENCODING, CharEncoding.UTF_8);
        Velocity.setProperty(Velocity.OUTPUT_ENCODING, CharEncoding.UTF_8);
        Velocity.init();
    }

    private Template template;

    public VelocityExcelBuilder() {
        widthStrategy(WidthStrategy.AUTO_WIDTH);
    }

    /**
     * 设置模板信息
     *
     * @param path 模板路径，相对路径
     */
    @Override
    public ExcelBuilder template(String path) {
        template = Velocity.getTemplate(path);
        return this;
    }

    @Override
    protected <T> void render(Map<String, T> data, Writer out) throws Exception {
        Objects.requireNonNull(template, "The template cannot be empty. Please set the template first.");
        VelocityContext context = new VelocityContext(data);
        template.merge(context, out);
    }
}

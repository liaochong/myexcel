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
import org.apache.commons.codec.CharEncoding;
import org.beetl.core.Configuration;
import org.beetl.core.GroupTemplate;
import org.beetl.core.ResourceLoader;
import org.beetl.core.Template;
import org.beetl.core.resource.ClasspathResourceLoader;
import org.beetl.core.resource.FileResourceLoader;

import java.io.IOException;
import java.io.Writer;
import java.nio.charset.StandardCharsets;
import java.util.HashMap;
import java.util.Map;
import java.util.Objects;
import java.util.function.Supplier;

/**
 * @author liaochong
 * @version 1.0
 */
public class BeetlTemplateHandler extends AbstractTemplateHandler<GroupTemplate, Template> {

    private static final Map<String, GroupTemplate> CFG_MAP = new HashMap<>();

    @Override
    protected void setTemplateEngine(String dirPath, Supplier<GroupTemplate> supplier, String fileName) {
        GroupTemplate groupTemplate = CFG_MAP.get(dirPath);
        if (groupTemplate == null) {
            groupTemplate = supplier.get();
        }
        templateEngine = groupTemplate.getTemplate(fileName);
    }

    @Override
    protected GroupTemplate getTemplateEngineSupplier(String dirPath) {
        synchronized (BeetlTemplateHandler.class) {
            GroupTemplate groupTemplate = CFG_MAP.get(dirPath);
            if (groupTemplate != null) {
                return groupTemplate;
            }
            ResourceLoader resourceLoader;
            if (Objects.equals(dirPath, CLASSPATH)) {
                resourceLoader = new ClasspathResourceLoader();
            } else {
                resourceLoader = new FileResourceLoader(dirPath, StandardCharsets.UTF_8.name());
            }
            Configuration cfg;
            try {
                cfg = Configuration.defaultConfiguration();
            } catch (IOException e) {
                throw new ExcelBuildException("Set Beetl configuration failure", e);
            }
            cfg.setCharset(CharEncoding.UTF_8);
            groupTemplate = new GroupTemplate(resourceLoader, cfg);
            CFG_MAP.put(dirPath, groupTemplate);
            return groupTemplate;
        }
    }

    @Override
    protected <F> void render(Map<String, F> renderData, Writer out) throws Exception {
        templateEngine.binding(renderData);
        templateEngine.renderTo(out);
    }
}

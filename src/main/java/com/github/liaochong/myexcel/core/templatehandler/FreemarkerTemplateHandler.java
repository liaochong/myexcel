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
import freemarker.template.Configuration;
import freemarker.template.Template;
import freemarker.template.TemplateExceptionHandler;
import org.apache.commons.codec.CharEncoding;

import java.io.File;
import java.io.IOException;
import java.io.Writer;
import java.util.HashMap;
import java.util.Map;
import java.util.Objects;
import java.util.function.Supplier;

/**
 * @author liaochong
 * @version 1.0
 */
public class FreemarkerTemplateHandler extends AbstractTemplateHandler<Configuration, Template> {

    protected static final Map<String, Configuration> CFG_MAP = new HashMap<>();

    @Override
    protected void setTemplateEngine(String dirPath, Supplier<Configuration> supplier, String fileName) {
        Configuration configuration = CFG_MAP.getOrDefault(dirPath, supplier.get());
        try {
            templateEngine = configuration.getTemplate(fileName);
        } catch (IOException e) {
            throw ExcelBuildException.of("Failed to get freemarker template", e);
        }
    }

    @Override
    protected Configuration getTemplateEngineSupplier(String dirPath) {
        synchronized (FreemarkerTemplateHandler.class) {
            Configuration configuration = CFG_MAP.get(dirPath);
            if (configuration != null) {
                return configuration;
            }
            configuration = new Configuration(Configuration.VERSION_2_3_23);
            configuration.setTemplateExceptionHandler(TemplateExceptionHandler.RETHROW_HANDLER);
            configuration.setDefaultEncoding(CharEncoding.UTF_8);
            try {
                if (Objects.equals(dirPath, CLASSPATH)) {
                    configuration.setClassLoaderForTemplateLoading(Thread.currentThread().getContextClassLoader(), "/");
                } else {
                    configuration.setDirectoryForTemplateLoading(new File(dirPath));
                }
            } catch (IOException e) {
                throw new ExcelBuildException("Set Freemarker directory failure", e);
            }
            CFG_MAP.put(dirPath, configuration);
            return configuration;
        }
    }

    @Override
    protected <E> void render(Map<String, E> renderData, Writer out) throws Exception {
        templateEngine.process(renderData, out);
    }
}

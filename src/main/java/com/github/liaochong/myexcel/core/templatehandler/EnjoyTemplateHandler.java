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

import com.jfinal.template.Engine;
import com.jfinal.template.Template;

import java.io.Writer;
import java.util.HashMap;
import java.util.Map;
import java.util.Objects;
import java.util.function.Supplier;

/**
 * @author liaochong
 * @version 1.0
 */
public class EnjoyTemplateHandler extends AbstractTemplateHandler<Engine, Template> {

    private static final Map<String, Engine> CFG_MAP = new HashMap<>();

    @Override
    protected void setTemplateEngine(String dirPath, Supplier<Engine> supplier, String fileName) {
        Engine engine = CFG_MAP.getOrDefault(dirPath, supplier.get());
        templateEngine = engine.getTemplate(fileName);
    }

    @Override
    protected Engine getTemplateEngineSupplier(String dirPath) {
        synchronized (EnjoyTemplateHandler.class) {
            Engine engine = CFG_MAP.get(dirPath);
            if (engine != null) {
                return engine;
            }
            engine = Engine.create("myexcel_" + dirPath);
            Engine.setFastMode(true);
            if (Objects.equals(dirPath, CLASSPATH)) {
                engine.setBaseTemplatePath(null);
                engine.setToClassPathSourceFactory();
            } else {
                engine.setBaseTemplatePath(dirPath);
            }
            CFG_MAP.put(dirPath, engine);
            return engine;
        }
    }

    @Override
    protected <F> void render(Map<String, F> renderData, Writer out) throws Exception {
        templateEngine.render(renderData, out);
    }
}

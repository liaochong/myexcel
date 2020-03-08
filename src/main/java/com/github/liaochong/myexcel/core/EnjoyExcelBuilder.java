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

import com.jfinal.template.Engine;
import com.jfinal.template.Template;

import java.io.Writer;
import java.util.HashMap;
import java.util.Map;
import java.util.Objects;
import java.util.function.Supplier;

/**
 * jfinal enjoy
 *
 * @author liaochong
 * @version 1.0
 */
public class EnjoyExcelBuilder extends AbstractExcelBuilder {

    private static final Map<String, Engine> CFG_MAP = new HashMap<>();

    private Template template;

    @Override
    public ExcelBuilder classpathTemplate(String path) {
        doSetEngine(CLASSPATH, () -> doGetEngine(CLASSPATH), path);
        return this;
    }

    @Override
    public ExcelBuilder fileTemplate(String dirPath, String fileName) {
        doSetEngine(dirPath, () -> doGetEngine(dirPath), fileName);
        return this;
    }

    private void doSetEngine(String dirPath, Supplier<Engine> supplier, String fileName) {
        Engine engine = CFG_MAP.get(dirPath);
        if (engine == null) {
            engine = supplier.get();
        }
        template = engine.getTemplate(fileName);
    }

    @Override
    protected <T> void render(Map<String, T> renderData, Writer out) throws Exception {
        checkTemplate(template);
        template.render(renderData, out);
    }

    private synchronized Engine doGetEngine(String dirPath) {
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

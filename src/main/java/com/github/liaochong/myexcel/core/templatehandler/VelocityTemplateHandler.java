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

import org.apache.commons.codec.CharEncoding;
import org.apache.velocity.Template;
import org.apache.velocity.VelocityContext;
import org.apache.velocity.app.Velocity;
import org.apache.velocity.runtime.resource.loader.ClasspathResourceLoader;

import java.io.Writer;
import java.util.HashMap;
import java.util.Map;
import java.util.function.Supplier;

/**
 * @author liaochong
 * @version 1.0
 */
public class VelocityTemplateHandler extends AbstractTemplateHandler<Template, Template> {

    static {
        Velocity.setProperty(Velocity.RESOURCE_LOADER, "classpath");
        Velocity.setProperty("classpath.resource.loader.class", ClasspathResourceLoader.class.getName());
        Velocity.setProperty(Velocity.ENCODING_DEFAULT, CharEncoding.UTF_8);
        Velocity.setProperty(Velocity.INPUT_ENCODING, CharEncoding.UTF_8);
        Velocity.init();
    }

    @Override
    public VelocityTemplateHandler classpathTemplate(String path) {
        templateEngine = Velocity.getTemplate(path);
        return this;
    }

    @Override
    public VelocityTemplateHandler fileTemplate(String dirPath, String fileName) {
        throw new UnsupportedOperationException();
    }

    @Override
    protected void setTemplateEngine(String dirPath, Supplier<Template> supplier, String fileName) {

    }

    @Override
    protected Template getTemplateEngineSupplier(String dirPath) {
        return null;
    }

    @Override
    protected <F> void render(Map<String, F> renderData, Writer out) throws Exception {
        VelocityContext context = new VelocityContext(new HashMap<>(renderData));
        templateEngine.merge(context, out);
    }
}

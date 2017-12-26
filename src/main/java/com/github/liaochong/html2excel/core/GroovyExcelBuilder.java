package com.github.liaochong.html2excel.core;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.io.OutputStreamWriter;
import java.io.Reader;
import java.io.Writer;
import java.util.Map;

import org.apache.commons.codec.CharEncoding;
import org.apache.poi.ss.usermodel.Workbook;

import groovy.lang.Writable;
import groovy.text.Template;
import groovy.text.markup.MarkupTemplateEngine;
import groovy.text.markup.TemplateConfiguration;

/**
 * @author liaochong
 * @version 1.0
 */
public class GroovyExcelBuilder extends ExcelBuilder {

    private Template template;

    @Override
    public ExcelBuilder getTemplate(String path) {
        TemplateConfiguration config = new TemplateConfiguration();
        MarkupTemplateEngine engine = new MarkupTemplateEngine(config);
        try {
            InputStream is = Thread.currentThread().getContextClassLoader().getResourceAsStream(path);
            Reader reader = new InputStreamReader(is);
            template = engine.createTemplate(reader);
            return this;
        } catch (ClassNotFoundException | IOException e) {
            throw new RuntimeException(e);
        }
    }

    @Override
    public Workbook build(Map<String, Object> renderData) {
        try {
            File htmlFile = this.createTempFile("groovy_temp_");
            Writer out = new BufferedWriter(new OutputStreamWriter(new FileOutputStream(htmlFile), CharEncoding.UTF_8));

            Writable output = template.make(renderData);
            output.writeTo(out);

            Workbook workbook = HtmlToExcelFactory.readHtml(htmlFile, htmlToExcelFactory).build();
            this.deleteTempFile(htmlFile);
            return workbook;
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }
}

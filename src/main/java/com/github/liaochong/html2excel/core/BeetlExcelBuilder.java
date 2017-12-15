package com.github.liaochong.html2excel.core;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStreamWriter;
import java.io.Writer;
import java.util.Map;

import org.apache.commons.codec.CharEncoding;
import org.apache.poi.ss.usermodel.Workbook;
import org.beetl.core.Configuration;
import org.beetl.core.GroupTemplate;
import org.beetl.core.Template;
import org.beetl.core.resource.ClasspathResourceLoader;

/**
 * beetl excel创建者
 *
 * @author liaochong
 * @version 1.0
 */
public class BeetlExcelBuilder extends ExcelBuilder {

    private Template template;

    @Override
    public ExcelBuilder getTemplate(String path) {
        try {
            String[] filePath = this.splitFilePath(path);
            ClasspathResourceLoader resourceLoader = new ClasspathResourceLoader(filePath[0]);
            Configuration cfg = Configuration.defaultConfiguration();
            cfg.setCharset(CharEncoding.UTF_8);
            GroupTemplate gt = new GroupTemplate(resourceLoader, cfg);
            template = gt.getTemplate(filePath[1]);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
        return this;
    }

    @Override
    public Workbook build(Map<String, Object> renderData) {
        try {
            File htmlFile = this.createTempFile("beetl_temp_");
            Writer out = new BufferedWriter(new OutputStreamWriter(new FileOutputStream(htmlFile), CharEncoding.UTF_8));

            template.binding(renderData);
            template.renderTo(out);
            Workbook workbook = HtmlToExcelFactory.readHtml(htmlFile, htmlToExcelFactory).build();
            this.deleteTempFile(htmlFile);
            return workbook;
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }
}

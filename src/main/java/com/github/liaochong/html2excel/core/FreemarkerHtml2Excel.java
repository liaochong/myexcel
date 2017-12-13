package com.github.liaochong.html2excel.core;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStreamWriter;
import java.io.Writer;
import java.util.Map;
import java.util.UUID;

import org.apache.commons.codec.CharEncoding;
import org.apache.poi.ss.usermodel.Workbook;

import freemarker.template.Configuration;
import freemarker.template.Template;
import freemarker.template.TemplateExceptionHandler;

/**
 * @author liaochong
 * @version 1.0
 */
public class FreemarkerHtml2Excel {

    private Template template;

    /**
     * 设置模板信息
     * 
     * @param basePackagePath 模板路径，相对路径
     * @param templateName 模板名称
     */
    public void getTemplate(String basePackagePath, String templateName) {
        Configuration cfg = new Configuration(Configuration.VERSION_2_3_23);
        cfg.setTemplateExceptionHandler(TemplateExceptionHandler.RETHROW_HANDLER);
        cfg.setDefaultEncoding(CharEncoding.UTF_8);
        cfg.setClassLoaderForTemplateLoading(Thread.currentThread().getContextClassLoader(), basePackagePath);
        try {
            template = cfg.getTemplate(templateName);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    /**
     * 构建
     * 
     * @param data 模板参数
     * @return Workbook
     */
    public Workbook build(Map<String, Object> data) {
        try {
            File htmlFile = File.createTempFile("temp" + UUID.randomUUID(), ".html");
            Writer out = new BufferedWriter(new OutputStreamWriter(new FileOutputStream(htmlFile), CharEncoding.UTF_8));
            template.process(data, out);
            Workbook workbook = Html2Excel.readHtml(htmlFile).build();
            boolean isDeleted = htmlFile.delete();
            if (!isDeleted) {
                throw new RuntimeException();
            }
            return workbook;
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }

}

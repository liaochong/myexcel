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
public class FreemarkerHtml2Excel implements Html2ExcelEnable {

    private Html2Excel html2Excel = new Html2Excel();

    private Template template;

    @Override
    public Html2ExcelEnable type(WorkbookType workbookType) {
        html2Excel.type(workbookType);
        return this;
    }

    @Override
    public Html2ExcelEnable useDefaultStyle() {
        html2Excel.useDefaultStyle();
        return this;
    }

    /**
     * 设置模板信息
     * 
     * @param path 模板路径，相对路径
     */
    @Override
    public Html2ExcelEnable getTemplate(String path) {
        Configuration cfg = new Configuration(Configuration.VERSION_2_3_23);
        cfg.setTemplateExceptionHandler(TemplateExceptionHandler.RETHROW_HANDLER);
        cfg.setDefaultEncoding(CharEncoding.UTF_8);

        int lastPackageIndex = path.lastIndexOf("/");
        String basePackagePath = path.substring(0, lastPackageIndex);
        cfg.setClassLoaderForTemplateLoading(Thread.currentThread().getContextClassLoader(), basePackagePath);
        try {
            String templateName = path.substring(lastPackageIndex + 1);
            template = cfg.getTemplate(templateName);
            return this;
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
    @Override
    public Workbook build(Map<String, Object> data) {
        try {
            File htmlFile = File.createTempFile("temp" + UUID.randomUUID(), ".html");
            Writer out = new BufferedWriter(new OutputStreamWriter(new FileOutputStream(htmlFile), CharEncoding.UTF_8));
            template.process(data, out);
            Workbook workbook = Html2Excel.readHtml(htmlFile, html2Excel).build();
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

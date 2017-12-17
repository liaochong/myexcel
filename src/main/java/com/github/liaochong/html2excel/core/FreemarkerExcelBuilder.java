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

import com.github.liaochong.html2excel.exception.ExcelBuildException;
import freemarker.template.Configuration;
import freemarker.template.Template;
import freemarker.template.TemplateExceptionHandler;

/**
 * freemarker的excel创建者
 *
 * @author liaochong
 * @version 1.0
 */
public class FreemarkerExcelBuilder extends ExcelBuilder {

    private Template template;

    /**
     * 设置模板信息
     * 
     * @param path 模板路径，相对路径
     */
    @Override
    public ExcelBuilder getTemplate(String path) {
        try {
            Configuration cfg = new Configuration(Configuration.VERSION_2_3_23);
            cfg.setTemplateExceptionHandler(TemplateExceptionHandler.RETHROW_HANDLER);
            cfg.setDefaultEncoding(CharEncoding.UTF_8);

            String[] filePath = this.splitFilePath(path);
            cfg.setClassLoaderForTemplateLoading(Thread.currentThread().getContextClassLoader(), filePath[0]);
            template = cfg.getTemplate(filePath[1]);
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
            File htmlFile = this.createTempFile("freemarker_temp_");
            Writer out = new BufferedWriter(new OutputStreamWriter(new FileOutputStream(htmlFile), CharEncoding.UTF_8));
            template.process(data, out);
            Workbook workbook = HtmlToExcelFactory.readHtml(htmlFile, htmlToExcelFactory).build();
            this.deleteTempFile(htmlFile);
            return workbook;
        } catch (Exception e) {
            throw ExcelBuildException.of("创建excel失败", e);
        }
    }

}

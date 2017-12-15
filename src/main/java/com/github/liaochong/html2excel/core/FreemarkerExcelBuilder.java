package com.github.liaochong.html2excel.core;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStreamWriter;
import java.io.Writer;
import java.util.Map;
import java.util.Objects;

import org.apache.commons.codec.CharEncoding;
import org.apache.poi.ss.usermodel.Workbook;

import freemarker.template.Configuration;
import freemarker.template.Template;
import freemarker.template.TemplateExceptionHandler;
import freemarker.template.Version;

/**
 * freemarker的excel创建者
 *
 * @author liaochong
 * @version 1.0
 */
public class FreemarkerExcelBuilder extends ExcelBuilder {

    private Template template;

    private Version version;

    /**
     * 增加版本设置，若未设置，默认采用版本 Configuration.VERSION_2_3_23
     * 
     * @param version 版本
     * @return FreemarkerExcelBuilder
     */
    public FreemarkerExcelBuilder version(Version version) {
        this.version = version;
        return this;
    }

    /**
     * 设置模板信息
     * 
     * @param path 模板路径，相对路径
     */
    @Override
    public ExcelBuilder getTemplate(String path) {
        if (Objects.isNull(version)) {
            version = Configuration.VERSION_2_3_23;
        }
        Configuration cfg = new Configuration(version);
        cfg.setTemplateExceptionHandler(TemplateExceptionHandler.RETHROW_HANDLER);
        cfg.setDefaultEncoding(CharEncoding.UTF_8);

        String[] filePath = this.splitFilePath(path);
        cfg.setClassLoaderForTemplateLoading(Thread.currentThread().getContextClassLoader(), filePath[0]);
        try {
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
            throw new RuntimeException(e);
        }
    }

}

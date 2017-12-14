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
public class BeetlExcelBuilder implements ExcelBuilder {

    private HtmlToExcelFactory htmlToExcelFactory = new HtmlToExcelFactory();

    private Template template;

    @Override
    public ExcelBuilder type(WorkbookType workbookType) {
        htmlToExcelFactory.type(workbookType);
        return this;
    }

    @Override
    public ExcelBuilder useDefaultStyle() {
        htmlToExcelFactory.useDefaultStyle();
        return this;
    }

    @Override
    public ExcelBuilder getTemplate(String path) {
        try {
            int lastPackageIndex = path.lastIndexOf("/");
            String basePackagePath = path.substring(0, lastPackageIndex);

            ClasspathResourceLoader resourceLoader = new ClasspathResourceLoader(basePackagePath);
            Configuration cfg = Configuration.defaultConfiguration();
            GroupTemplate gt = new GroupTemplate(resourceLoader, cfg);
            template = gt.getTemplate(path.substring(lastPackageIndex));
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
        return this;
    }

    @Override
    public Workbook build(Map<String, Object> renderData) {
        try {
            template.binding(renderData);
            File htmlFile = File.createTempFile("beetl_temp_" + UUID.randomUUID(), ".html");
            Writer out = new BufferedWriter(new OutputStreamWriter(new FileOutputStream(htmlFile), CharEncoding.UTF_8));
            template.renderTo(out);
            Workbook workbook = HtmlToExcelFactory.readHtml(htmlFile, htmlToExcelFactory).build();
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

package com.github.liaochong.myexcel.core;

import com.github.liaochong.myexcel.core.templatehandler.ThymeleafTemplateHandler;
import com.github.liaochong.myexcel.utils.FileExportUtil;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.jupiter.api.Test;

import java.io.File;
import java.util.Collections;

/**
 * @author liaochong
 * @version 1.0
 */
class TemplateStreamExcelBuilderTest extends BasicTest {

    @Test
    public void test() throws Exception {
        TemplateStreamExcelBuilder templateStreamExcelBuilder = TemplateStreamExcelBuilder.of(ThymeleafTemplateHandler.class);
        templateStreamExcelBuilder.append("/templates/thymeleafToExcelExample.html", Collections.emptyMap());
        Workbook workbook = templateStreamExcelBuilder.build();
        FileExportUtil.export(workbook, new File(TEST_OUTPUT_DIR + "template_stream_thymeleaf_build.xlsx"));
    }

}
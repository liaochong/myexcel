package com.github.liaochong.example.controller;

import com.github.liaochong.example.pojo.ArtCrowd;
import com.github.liaochong.myexcel.core.DefaultExcelBuilder;
import com.github.liaochong.myexcel.core.strategy.WidthStrategy;
import com.github.liaochong.myexcel.utils.AttachmentExportUtil;
import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.web.bind.annotation.GetMapping;

import javax.servlet.http.HttpServletResponse;
import java.util.List;

public class Style {
    private final DefaultExcelBuilderExampleController defaultExcelBuilderExampleController;

    public Style(DefaultExcelBuilderExampleController defaultExcelBuilderExampleController) {
        this.defaultExcelBuilderExampleController = defaultExcelBuilderExampleController;
    }

    @GetMapping("/default/noStyle/example")
    public void defaultBuildWithNoStyle(HttpServletResponse response) throws Exception {
        List<ArtCrowd> dataList = defaultExcelBuilderExampleController.getDataList();
        Workbook workbook = DefaultExcelBuilder.of(ArtCrowd.class).noStyle().build(dataList);
        AttachmentExportUtil.export(workbook, "艺术生信息", response);
    }

    @GetMapping("/default/autoWidth/example")
    public void defaultBuildWithAutoWidth(HttpServletResponse response) throws Exception {
        List<ArtCrowd> dataList = defaultExcelBuilderExampleController.getDataList();
        Workbook workbook = DefaultExcelBuilder.of(ArtCrowd.class).widthStrategy(WidthStrategy.AUTO_WIDTH).build(dataList);
        AttachmentExportUtil.export(workbook, "艺术生信息", response);
    }

    @GetMapping("/default/continue/example")
    public void defaultBuildWithWorkbook(HttpServletResponse response) throws Exception {
        List<ArtCrowd> dataList = defaultExcelBuilderExampleController.getDataList();
        Workbook workbook = DefaultExcelBuilder.of(ArtCrowd.class).widthStrategy(WidthStrategy.AUTO_WIDTH).build(dataList);

        dataList = defaultExcelBuilderExampleController.getDataList();
        workbook = DefaultExcelBuilder.of(ArtCrowd.class, workbook).sheetName("sheet2").widthStrategy(WidthStrategy.NO_AUTO).build(dataList);
        AttachmentExportUtil.export(workbook, "艺术生信息", response);
    }
}
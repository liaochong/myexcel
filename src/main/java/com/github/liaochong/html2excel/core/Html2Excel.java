/*
 * Copyright 2017 the original author or authors.
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *      http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
package com.github.liaochong.html2excel.core;

import java.io.File;
import java.util.ArrayList;
import java.util.EnumMap;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.Optional;
import java.util.function.Predicate;
import java.util.stream.Collectors;

import org.apache.commons.codec.CharEncoding;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;

import com.github.liaochong.html2excel.core.style.CellStyleFactory;
import com.github.liaochong.html2excel.core.style.TdCellStyle;
import com.github.liaochong.html2excel.core.style.ThCellStyle;
import com.github.liaochong.html2excel.exception.NoTablesException;
import com.github.liaochong.html2excel.utils.TdUtils;

/**
 * Html2Excel
 * <p>
 * 用于将html table解析成excel
 * </p>
 *
 * @author liaochong
 * @version 1.0
 */
public class Html2Excel {
    /**
     * html解析后文档
     */
    private Document document;
    /**
     * excel workbook
     */
    private Workbook workbook;
    /**
     * tr容器
     */
    private List<Tr> trContainer;
    /**
     * 最大列数
     */
    private int maxCols;
    /**
     * 是否使用默认样式
     */
    private boolean useDefaultStyle;
    /**
     * 样式容器
     */
    private Map<Tag, CellStyle> cellStyleFactoryEnumMap;
    /**
     * 每列最大宽度
     */
    private Map<Integer, Integer> colMaxWidthMap;

    private Html2Excel(Document document) {
        this.document = document;
    }

    /**
     * 读取html
     * 
     * @param htmlFile html文件
     * @throws Exception 解析异常
     */
    public static Html2Excel readHtml(File htmlFile) throws Exception {
        Document document = Jsoup.parse(htmlFile, CharEncoding.UTF_8);
        return new Html2Excel(document);
    }

    /**
     * 添加标题样式
     * 
     * @param cellStyleFactory 样式工厂
     * @return Html2Excel
     */
    public Html2Excel addThStyle(CellStyleFactory cellStyleFactory) {
        if (Objects.isNull(cellStyleFactoryEnumMap)) {
            cellStyleFactoryEnumMap = new EnumMap<>(Tag.class);
        }
        cellStyleFactoryEnumMap.put(Tag.th, cellStyleFactory.supply(workbook));
        return this;
    }

    /**
     * 添加单元格样式
     * 
     * @param cellStyleFactory 样式工厂
     * @return Html2Excel
     */
    public Html2Excel addTdStyle(CellStyleFactory cellStyleFactory) {
        if (Objects.isNull(cellStyleFactoryEnumMap)) {
            cellStyleFactoryEnumMap = new EnumMap<>(Tag.class);
        }
        cellStyleFactoryEnumMap.put(Tag.td, cellStyleFactory.supply(workbook));
        return this;
    }

    /**
     * 设置使用默认样式
     * 
     * @return Html2Excel
     */
    public Html2Excel useDefaultStyle() {
        useDefaultStyle = true;
        return this;
    }

    /**
     * 开始解析
     * 
     * @return Workbook
     */
    public Workbook parse() {
        Elements tables = document.getElementsByTag(Tag.table.name());
        workbook = new XSSFWorkbook();
        if (tables.isEmpty()) {
            throw NoTablesException.of("There is no any table exist");
        }
        // 使用默认样式时，加载默认样式
        if (useDefaultStyle) {
            if (Objects.isNull(cellStyleFactoryEnumMap)) {
                cellStyleFactoryEnumMap = new EnumMap<>(Tag.class);
            }
            cellStyleFactoryEnumMap.putIfAbsent(Tag.th, new ThCellStyle().supply(workbook));
            cellStyleFactoryEnumMap.putIfAbsent(Tag.td, new TdCellStyle().supply(workbook));
        }
        for (int i = 0; i < tables.size(); i++) {
            this.processTable(tables.get(i), i);
        }
        return workbook;
    }

    /**
     * 解析每一个table
     * 
     * @param table 表格
     */
    private void processTable(Element table, int index) {
        this.initialize();
        Elements trs = table.getElementsByTag(Tag.tr.name());
        for (int i = 0; i < trs.size(); i++) {
            Tr tr = new Tr(i);
            trContainer.add(tr);
            this.processTr(trs.get(i), tr);
        }
        List<Td> allTds = this.adjust();
        Sheet sheet = this.getSheet(table, index);
        Predicate<Td> predicate = td -> td.getRowSpan() > 0 || td.getColSpan() > 0;
        allTds.stream().filter(predicate).forEach(td -> sheet.addMergedRegion(new CellRangeAddress(td.getX(),
                TdUtils.get(td::getRowSpan, td::getX), td.getY(), TdUtils.get(td::getColSpan, td::getY))));

        this.setColMaxWidthMap(allTds, sheet);
        allTds.forEach(td -> this.setCell(td, sheet));
    }

    /**
     * 初始化，每解析一个表格需要重新初始化
     */
    private void initialize() {
        maxCols = 0;
        trContainer = new ArrayList<>();
    }

    /**
     * 设置每列最大宽度
     * 
     * @param allTds 所有单元格
     */
    private void setColMaxWidthMap(List<Td> allTds, Sheet sheet) {
        colMaxWidthMap = new HashMap<>(maxCols);
        allTds.parallelStream().forEach(td -> {
            int width = TdUtils.getStringWidth(td.getContent());
            Integer maxWidth = colMaxWidthMap.get(td.getY());
            if (Objects.isNull(maxWidth) || maxWidth < width) {
                colMaxWidthMap.put(td.getY(), width);
            }
        });
        colMaxWidthMap.forEach((key, value) -> {
            sheet.setColumnWidth(key, value * 2 * 255);
        });
    }

    /**
     * 设置单元格
     * 
     * @param td 单元格
     * @param sheet 单元格所在的sheet
     */
    private void setCell(Td td, Sheet sheet) {
        Cell cell = sheet.getRow(td.getX()).getCell(td.getY());
        cell.setCellValue(td.getContent());
        if (useDefaultStyle) {
            if (td.isTh()) {
                cell.setCellStyle(cellStyleFactoryEnumMap.get(Tag.th));
            } else {
                cell.setCellStyle(cellStyleFactoryEnumMap.get(Tag.td));
            }
        }
    }

    /**
     * 获取sheet
     *
     * @param table 表格
     * @param index 索引
     * @return Sheet
     */
    private Sheet getSheet(final Element table, int index) {
        Elements captions = table.getElementsByTag(Tag.caption.name());
        String sheetName;
        if (!captions.isEmpty()) {
            sheetName = captions.first().text();
        } else {
            sheetName = "sheet" + ++index;
        }
        Sheet sheet = workbook.createSheet(sheetName);
        // 创建空白单元格
        for (int i = 0; i < trContainer.size(); i++) {
            Row row = sheet.createRow(i);
            for (int j = 0; j <= maxCols; j++) {
                row.createCell(j);
            }
        }
        return sheet;
    }

    /**
     * 处理行元素
     *
     * @param tr tr
     * @param container tr容器
     */
    private void processTr(Element tr, Tr container) {
        Elements ths = tr.getElementsByTag(Tag.th.name());
        this.processing(ths, container, true);
        Elements tds = tr.getElementsByTag(Tag.td.name());
        this.processing(tds, container, false);
    }

    /**
     * 处理行内元素
     *
     * @param elements 元素：th、td
     * @param container 元素容器
     * @param isTh 是否为表格标题
     */
    private void processing(Elements elements, Tr container, boolean isTh) {
        if (elements.isEmpty()) {
            return;
        }
        for (int i = 0; i < elements.size(); i++) {
            Td td = new Td();
            td.setTh(isTh);
            td.setX(container.getIndex());
            td.setY(i);

            Element element = elements.get(i);
            String colSpan = element.attr(Tag.colspan.name());
            if (StringUtils.isNotBlank(colSpan)) {
                td.setColSpan(Integer.parseInt(colSpan));
            }
            String rowSpan = element.attr(Tag.rowspan.name());
            if (StringUtils.isNotBlank(rowSpan)) {
                td.setRowSpan(Integer.parseInt(rowSpan));
            }
            td.setContent(element.text());
            container.getTds().add(td);
        }
    }

    /**
     * 调整表格单元格位置
     *
     * @return 所有单元格
     */
    private List<Td> adjust() {
        List<Td> allTds = trContainer.stream().flatMap(tr -> tr.getTds().stream()).collect(Collectors.toList());
        // 排除第一行
        Predicate<Tr> predicate = tr -> tr.getIndex() > 0;
        trContainer.stream().filter(predicate)
                .forEach(tr -> tr.getTds().parallelStream().forEach(td -> this.adjust(allTds, td)));
        maxCols = allTds.stream().mapToInt(Td::getY).max().orElseThrow(() -> new NoTablesException("不存在任何单元格"));
        return allTds;
    }

    /**
     * 调整表格单元格位置
     * 
     * @param allTds 所有单元格
     * @param td 当前单元格
     */
    private void adjust(List<Td> allTds, Td td) {
        Predicate<Td> predicate = prevTd -> prevTd.getX() < td.getX() && prevTd.getY() == td.getY()
                && TdUtils.get(prevTd::getRowSpan, prevTd::getX) >= td.getX();
        Optional<Td> findResult = allTds.stream().filter(predicate).findAny();
        if (!findResult.isPresent()) {
            return;
        }
        Td sameColTd = findResult.get();
        int prevTdColSpan = sameColTd.getColSpan();
        int realY = prevTdColSpan > 0 ? td.getY() + prevTdColSpan : td.getY() + 1;
        td.setY(realY);
    }

    private enum Tag {
        /**
         * table
         */
        table,
        /**
         * caption
         */
        caption,
        /**
         * tr
         */
        tr,
        /**
         * th
         */
        th,
        /**
         * td
         */
        td,
        /**
         * colspan
         */
        colspan,
        /**
         * rowspan
         */
        rowspan;
    }
}

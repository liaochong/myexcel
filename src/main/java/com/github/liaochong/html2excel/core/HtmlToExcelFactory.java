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
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.concurrent.CompletableFuture;
import java.util.concurrent.ConcurrentHashMap;
import java.util.function.Predicate;
import java.util.function.ToIntFunction;
import java.util.stream.Collectors;

import org.apache.commons.codec.CharEncoding;
import org.apache.commons.collections4.CollectionUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
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

import com.github.liaochong.html2excel.core.style.TdCellStyle;
import com.github.liaochong.html2excel.core.style.ThCellStyle;
import com.github.liaochong.html2excel.exception.NoTablesException;
import com.github.liaochong.html2excel.utils.TdUtils;

/**
 * HtmlToExcelFactory
 * <p>
 * 用于将html table解析成excel
 * </p>
 *
 * @author liaochong
 * @version 1.0
 */
public class HtmlToExcelFactory {
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
     * 总列数
     */
    private int totalCols;
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
    /**
     * future
     */
    private CompletableFuture<Void> workbookFuture;
    /**
     * sheet容器
     */
    private Map<Integer, Sheet> sheetMap;

    public HtmlToExcelFactory() {
    }

    /**
     * 读取html
     * 
     * @param htmlFile html文件
     * @throws Exception 解析异常
     */
    public static HtmlToExcelFactory readHtml(File htmlFile) throws Exception {
        HtmlToExcelFactory factory = new HtmlToExcelFactory();
        factory.document = Jsoup.parse(htmlFile, CharEncoding.UTF_8);
        return factory;
    }

    /**
     * 读取html
     * 
     * @param htmlFile html文件
     * @param htmlToExcelFactory 实例对象
     * @return HtmlToExcelFactory
     * @throws Exception 解析异常
     */
    public static HtmlToExcelFactory readHtml(File htmlFile, HtmlToExcelFactory htmlToExcelFactory) throws Exception {
        htmlToExcelFactory.document = Jsoup.parse(htmlFile, CharEncoding.UTF_8);
        return htmlToExcelFactory;
    }

    /**
     * 设置使用默认样式
     * 
     * @return HtmlToExcelFactory
     */
    public HtmlToExcelFactory useDefaultStyle() {
        useDefaultStyle = true;
        return this;
    }

    /**
     * 设置workbook类型
     * 
     * @param workbookType 工作簿类型
     * @return HtmlToExcelFactory
     */
    public HtmlToExcelFactory type(WorkbookType workbookType) {
        if (WorkbookType.isXls(workbookType)) {
            workbook = new HSSFWorkbook();
        } else {
            workbook = new XSSFWorkbook();
        }
        return this;
    }

    /**
     * 开始解析
     * 
     * @return Workbook
     */
    public Workbook build() {
        Elements tables = document.getElementsByTag(Tag.table.name());
        if (tables.isEmpty()) {
            throw NoTablesException.of("There is no any table exist");
        }
        // 1、创建工作簿
        createWorkbook(tables);
        for (int i = 0; i < tables.size(); i++) {
            // 2、处理解析表格
            List<Td> tds = this.processTable(tables.get(i));
            // 3、设置单元格
            this.setUp(i, tds);
        }
        return workbook;
    }

    /**
     * 创建workbook，因为创建workbook比较耗时，异步处理
     * 
     * @param tables 表格集合
     */
    private void createWorkbook(Elements tables) {
        workbookFuture = CompletableFuture.runAsync(() -> {
            if (Objects.isNull(workbook)) {
                workbook = new XSSFWorkbook();
            }
            // 使用默认样式时，加载默认样式
            if (useDefaultStyle) {
                if (Objects.isNull(cellStyleFactoryEnumMap)) {
                    cellStyleFactoryEnumMap = new EnumMap<>(Tag.class);
                }
                cellStyleFactoryEnumMap.putIfAbsent(Tag.th, new ThCellStyle().supply(workbook));
                cellStyleFactoryEnumMap.putIfAbsent(Tag.td, new TdCellStyle().supply(workbook));
            }
            sheetMap = new ConcurrentHashMap<>(tables.size());
            for (int i = 0; i < tables.size(); i++) {
                Element table = tables.get(i);
                Elements captions = table.getElementsByTag(Tag.caption.name());
                String sheetName;
                if (!captions.isEmpty()) {
                    sheetName = captions.first().text();
                } else {
                    sheetName = "sheet" + (i + 1);
                }
                Sheet sheet = workbook.createSheet(sheetName);
                sheetMap.put(i, sheet);

                Elements trs = table.getElementsByTag(Tag.tr.name());
                int cloNum = trs.parallelStream().mapToInt(tr -> {
                    Elements tds = tr.children();
                    return tds.stream().mapToInt(td -> {
                        String colSpan = td.attr(Tag.colspan.name());
                        return StringUtils.isNotBlank(colSpan) ? Integer.parseInt(colSpan) : 1;
                    }).sum();
                }).max().orElse(0);

                for (int j = 0; j < trs.size(); j++) {
                    Row row = sheet.createRow(j);
                    for (int k = 0; k <= cloNum; k++) {
                        row.createCell(k);
                    }
                }
            }
        });
    }

    /**
     * 解析每一个table
     *
     * @param table 表格
     */
    private List<Td> processTable(Element table) {
        this.initialize();
        Elements trs = table.getElementsByTag(Tag.tr.name());
        for (int i = 0; i < trs.size(); i++) {
            Tr tr = new Tr(i);
            trContainer.add(tr);
            this.processTr(trs.get(i), tr);
        }
        this.getTotalCols();
        return this.adjust();
    }

    /**
     * 创建
     * 
     * @param tableIndex 表格索引
     */
    private void setUp(int tableIndex, List<Td> allTds) {
        workbookFuture.join();
        Sheet sheet = sheetMap.get(tableIndex);
        allTds.forEach(td -> this.setCell(td, sheet));
        colMaxWidthMap.forEach((key, value) -> sheet.setColumnWidth(key, value << 9));
    }

    /**
     * 初始化，每解析一个表格需要重新初始化
     */
    private void initialize() {
        totalCols = 0;
        trContainer = new ArrayList<>();
        colMaxWidthMap = new ConcurrentHashMap<>();
    }

    /**
     * 获取总列数
     */
    private void getTotalCols() {
        ToIntFunction<Tr> function = tr -> tr.getTds().stream().mapToInt(td -> TdUtils.get(td::getColSpan, td::getCol))
                .max().orElse(0);
        totalCols = trContainer.parallelStream().mapToInt(function).max()
                .orElseThrow(() -> new NoTablesException("不存在任何单元格"));
    }

    /**
     * 设置单元格
     *
     * @param td 单元格
     * @param sheet 单元格所在的sheet
     */
    private void setCell(Td td, Sheet sheet) {
        Cell cell = sheet.getRow(td.getRow()).getCell(td.getCol());
        cell.setCellValue(td.getContent());

        // 设置单元格样式
        int boundCol = TdUtils.get(td::getColSpan, td::getCol);
        int boundRow = TdUtils.get(td::getRowSpan, td::getRow);

        for (int i = td.getRow(); i <= boundRow; i++) {
            for (int j = td.getCol(); j <= boundCol; j++) {
                cell = sheet.getRow(i).getCell(j);
                if (useDefaultStyle) {
                    if (td.isTh()) {
                        cell.setCellStyle(cellStyleFactoryEnumMap.get(Tag.th));
                    } else {
                        cell.setCellStyle(cellStyleFactoryEnumMap.get(Tag.td));
                    }
                }
            }
        }
        if (td.getColSpan() > 0 || td.getRowSpan() > 0) {
            sheet.addMergedRegion(new CellRangeAddress(td.getRow(), boundRow, td.getCol(), boundCol));
        }
    }

    /**
     * 处理行元素
     *
     * @param tr tr
     * @param container tr容器
     */
    private void processTr(Element tr, Tr container) {
        Elements ths = tr.getElementsByTag(Tag.th.name());
        this.processTd(ths, container, true);
        Elements tds = tr.getElementsByTag(Tag.td.name());
        this.processTd(tds, container, false);
    }

    /**
     * 处理行内元素
     *
     * @param elements 元素：th、td
     * @param container 元素容器
     * @param isTh 是否为表格标题
     */
    private void processTd(Elements elements, Tr container, boolean isTh) {
        if (elements.isEmpty()) {
            return;
        }
        for (int i = 0; i < elements.size(); i++) {
            Td td = new Td();
            td.setTh(isTh);
            td.setRow(container.getIndex());
            // 除每行第一个单元格外，修正含跨列的单元格位置
            if (i > 0) {
                int shift = container.getTds().stream().filter(t -> t.getColSpan() > 0)
                        .mapToInt(t -> t.getColSpan() - 1).sum();
                td.setCol(i + shift);
            } else {
                td.setCol(i);
            }
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
            // 设置每列最宽宽度
            int width = TdUtils.getStringWidth(td.getContent());
            Integer maxWidth = colMaxWidthMap.get(td.getCol());
            if (Objects.isNull(maxWidth) || maxWidth < width) {
                colMaxWidthMap.put(td.getCol(), width);
            }
        }
    }

    /**
     * 调整表格单元格位置
     *
     * @return 所有单元格
     */
    private List<Td> adjust() {
        // 排除第一行
        for (int i = 1; i < trContainer.size(); i++) {
            int index = i;
            trContainer.get(i).getTds().forEach(td -> this.adjust(td, index));
        }
        return trContainer.stream().flatMap(tr -> tr.getTds().stream()).collect(Collectors.toList());
    }

    /**
     * 调整表格单元格位置
     *
     * @param td 单元格
     * @param trIndex 单元格所在行索引
     */
    private void adjust(Td td, int trIndex) {
        Predicate<Tr> predicate = tr -> tr.getIndex() < trIndex;
        List<Td> tds = trContainer.stream().filter(predicate).flatMap(tr -> tr.getTds().stream())
                .collect(Collectors.toList());
        if (CollectionUtils.isEmpty(tds)) {
            return;
        }
        for (int i = 0; i < totalCols; i++) {
            Td td1 = tds.stream().filter(prevTd -> prevTd.getCol() <= td.getCol()
                    && TdUtils.get(prevTd::getRowSpan, prevTd::getRow) >= td.getRow()).findFirst().orElse(null);
            if (Objects.isNull(td1)) {
                return;
            }
            int prevTdColSpan = td1.getColSpan();
            int realCol = prevTdColSpan > 0 ? td.getCol() + prevTdColSpan : td.getCol() + 1;
            td.setCol(realCol);
            tds.remove(td1);
        }
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

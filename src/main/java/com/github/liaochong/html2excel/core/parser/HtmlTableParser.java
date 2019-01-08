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
package com.github.liaochong.html2excel.core.parser;

import com.github.liaochong.html2excel.utils.StyleUtil;
import com.github.liaochong.html2excel.utils.TdUtil;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.codec.CharEncoding;
import org.apache.commons.collections4.CollectionUtils;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;

import java.io.File;
import java.io.IOException;
import java.util.Collections;
import java.util.Comparator;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.concurrent.ConcurrentHashMap;
import java.util.stream.Collectors;
import java.util.stream.IntStream;

/**
 * html table parser
 *
 * @author liaochong
 * @version 1.0
 */
@Slf4j
public class HtmlTableParser {

    /**
     * html解析后文档
     */
    private Document document;

    private HtmlTableParser() {

    }

    public static HtmlTableParser of(File htmlFile) throws IOException {
        Objects.requireNonNull(htmlFile);
        HtmlTableParser parser = new HtmlTableParser();
        parser.document = Jsoup.parse(htmlFile, CharEncoding.UTF_8);
        return parser;
    }

    public static HtmlTableParser of(String html) {
        Objects.requireNonNull(html);
        HtmlTableParser parser = new HtmlTableParser();
        parser.document = Jsoup.parse(html, CharEncoding.UTF_8);
        return parser;
    }

    /**
     * 获取所有表格
     *
     * @return 所有表格
     */
    public List<Table> getAllTable() {
        log.info("Start parsing html file");
        long startTime = System.currentTimeMillis();
        Elements tableElements = document.getElementsByTag(TableTag.table.name());
        List<Table> result = tableElements.stream().map(tableElement -> {
            Table table = new Table();
            Elements captionElements = tableElement.getElementsByTag(TableTag.caption.name());
            if (!captionElements.isEmpty()) {
                table.setCaption(captionElements.first().text());
            }
            this.parseTrOfTable(table, tableElement, StyleUtil.parseStyle(tableElement));
            return table;
        }).collect(Collectors.toList());
        log.info("Complete html file parsing,takes {} ms", System.currentTimeMillis() - startTime);
        return result;
    }

    /**
     * 解析table中的tr
     *
     * @param table table
     */
    private void parseTrOfTable(Table table, Element tableElement, Map<String, String> tableStyle) {
        List<Tr> trList = this.getTrList(tableElement, tableStyle);
        table.setTrList(trList);
        if (trList.isEmpty()) {
            return;
        }

        // 调整td位置,排除第一行，第一行不需要进行调整
        if (trList.size() == 1) {
            return;
        }
        trList.subList(1, trList.size()).parallelStream().forEach(tr -> {
            tr.getTdList().parallelStream().forEach(td -> this.adjustTdPosition(td, trList));
        });
    }

    /**
     * 获取Tr集合
     *
     * @return trList
     */
    private List<Tr> getTrList(Element tableElement, Map<String, String> tableStyle) {
        Map<Element, Map<String, String>> parentStyleMap = new ConcurrentHashMap<>();

        Elements trElements = tableElement.getElementsByTag(TableTag.tr.name());
        List<Tr> trList = IntStream.range(0, trElements.size()).parallel().mapToObj(index -> {
            Element trElement = trElements.get(index);
            Element parent = trElement.parent();
            Map<String, String> upperStyle;
            if (Objects.equals(parent, tableElement)) {
                upperStyle = tableStyle;
            } else {
                if (parentStyleMap.containsKey(parent)) {
                    upperStyle = parentStyleMap.get(parent);
                } else {
                    upperStyle = StyleUtil.mixStyle(tableStyle, StyleUtil.parseStyle(parent));
                    parentStyleMap.putIfAbsent(parent, upperStyle);
                }
            }
            Tr tr = new Tr(index);
            this.parseTdOfTr(tr, trElement, StyleUtil.mixStyle(upperStyle, StyleUtil.parseStyle(trElement)));
            return tr;
        }).collect(Collectors.toList());

        return trList.stream().sorted(Comparator.comparing(Tr::getIndex)).collect(Collectors.toList());
    }

    /**
     * 获取tr中的td
     *
     * @param tr tr
     */
    private void parseTdOfTr(Tr tr, Element trElement, Map<String, String> trStyle) {
        Elements tdElements = trElement.children();
        if (tdElements.isEmpty()) {
            tr.setColWidthMap(Collections.emptyMap());
            return;
        }
        tr.setColWidthMap(new HashMap<>(tdElements.size()));
        // 单元格偏移量
        int shift = 0;
        for (int i = 0, size = tdElements.size(); i < size; i++) {
            Element tdElement = tdElements.get(i);
            Td td = new Td();
            td.setContent(tdElement.text());
            td.setTh(Objects.equals(TableTag.th.name(), tdElement.tagName()));
            td.setRow(tr.getIndex());
            td.setStyle(StyleUtil.mixStyle(trStyle, StyleUtil.parseStyle(tdElement)));
            // 除每行第一个单元格外，修正含跨列的单元格位置
            td.setCol(i + shift);

            String colSpan = tdElement.attr(TableTag.colspan.name());
            td.setColSpan(TdUtil.getSpan(colSpan));

            String rowSpan = tdElement.attr(TableTag.rowspan.name());
            td.setRowSpan(TdUtil.getSpan(rowSpan));

            int rowBound = TdUtil.get(td::getRowSpan, td::getRow);
            td.setRowBound(rowBound);

            int colBound = TdUtil.get(td::getColSpan, td::getCol);
            td.setColBound(colBound);

            if (td.getColSpan() > 0) {
                shift += td.getColSpan() - 1;
            }

            tr.getTdList().add(td);

            // 设置每列宽度
            int width = TdUtil.getStringWidth(td.getContent());
            tr.getColWidthMap().put(td.getCol(), width);
        }
    }

    /**
     * 调整表格单元格位置
     *
     * @param td     单元格
     * @param trList 所有行
     */
    private void adjustTdPosition(Td td, List<Tr> trList) {
        List<Td> rowSpanTds = trList.subList(0, td.getRow()).stream()
                .flatMap(tr -> tr.getTdList().stream())
                .filter(t -> t.getRowSpan() > 0 && t.getCol() <= td.getCol()
                        && t.getRowBound() >= td.getRow())
                .collect(Collectors.toList());

        if (CollectionUtils.isEmpty(rowSpanTds)) {
            return;
        }
        rowSpanTds.forEach(t -> {
            int prevTdColSpan = t.getColSpan();
            int realCol = prevTdColSpan > 0 ? td.getCol() + prevTdColSpan : td.getCol() + 1;
            td.setCol(realCol);
        });

        // 重调
        int colBound = TdUtil.get(td::getColSpan, td::getCol);
        td.setColBound(colBound);
    }

    public enum TableTag {
        /**
         * table
         */
        table,
        /**
         * caption
         */
        caption,
        /**
         * thead
         */
        thead,
        /**
         * tbody
         */
        tbody,
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

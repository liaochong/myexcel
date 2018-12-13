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

import com.github.liaochong.html2excel.utils.StyleUtils;
import com.github.liaochong.html2excel.utils.TdUtils;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.codec.CharEncoding;
import org.apache.commons.collections4.CollectionUtils;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
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

    /**
     * 获取所有表格
     *
     * @return 所有表格
     */
    public List<Table> getAllTable() {
        log.info("Start parsing html file");
        long startTime = System.currentTimeMillis();
        Elements tableElements = document.getElementsByTag(TableTag.table.name());
        List<Table> result = IntStream.range(0, tableElements.size()).mapToObj(i -> {
            Element tableElement = tableElements.get(i);
            Table table = new Table();
            table.setIndex(i);
            table.setElement(tableElement);

            Elements captionElements = tableElement.getElementsByTag(TableTag.caption.name());
            if (!captionElements.isEmpty()) {
                table.setCaption(captionElements.first().text());
            }
            table.setStyleMap(StyleUtils.parseStyle(tableElement));

            this.parseTrOfTable(table);
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
    private void parseTrOfTable(Table table) {
        List<Tr> sortedTrList = this.getSortedTrList(table);
        table.setTrList(sortedTrList);
        if (sortedTrList.isEmpty()) {
            table.setColMaxWidthMap(Collections.emptyMap());
            return;
        }

        int lastColumnNum = sortedTrList.parallelStream().max(Comparator.comparing(Tr::getLastColumnNum)).get().getLastColumnNum();
        table.setLastColumnNum(lastColumnNum);

        this.setColMaxWidthMap(table);

        // 调整td位置,排除第一行，第一行不需要进行调整
        if (sortedTrList.size() == 1) {
            return;
        }
        sortedTrList.subList(1, sortedTrList.size()).parallelStream().forEach(tr -> {
            tr.getTdList().parallelStream().forEach(td -> this.adjustTdPosition(td, tr.getIndex(), sortedTrList));
        });
    }

    /**
     * 获取已排序的Tr集合
     *
     * @param table table
     * @return trList
     */
    private List<Tr> getSortedTrList(Table table) {
        Map<Element, Map<String, String>> parentStyleMap = new ConcurrentHashMap<>();

        Elements trElements = table.getElement().getElementsByTag(TableTag.tr.name());
        List<Tr> trList = IntStream.range(0, trElements.size()).parallel().mapToObj(index -> {
            Element trElement = trElements.get(index);
            Element parent = trElement.parent();
            Map<String, String> upperStyle;
            if (Objects.equals(parent, table.getElement())) {
                upperStyle = table.getStyleMap();
            } else {
                if (parentStyleMap.containsKey(parent)) {
                    upperStyle = parentStyleMap.get(parent);
                } else {
                    upperStyle = StyleUtils.mixStyle(table.getStyleMap(), StyleUtils.parseStyle(parent));
                    parentStyleMap.putIfAbsent(parent, upperStyle);
                }
            }
            Tr tr = new Tr(index);
            tr.setElement(trElement);
            tr.setStyle(StyleUtils.mixStyle(upperStyle, StyleUtils.parseStyle(trElement)));
            this.parseTdOfTr(tr);
            return tr;
        }).collect(Collectors.toList());

        // 重排序
        return trList.stream().sorted(Comparator.comparing(Tr::getIndex)).collect(Collectors.toList());
    }

    /**
     * 设置每列最大宽度
     *
     * @param table table
     */
    private void setColMaxWidthMap(Table table) {
        Map<Integer, Integer> colMaxWidthMap = new HashMap<>(table.getLastColumnNum());
        table.getTrList().stream().map(Tr::getColWidthMap).forEach(map -> {
            map.forEach((k, v) -> {
                Integer width = colMaxWidthMap.get(k);
                if (Objects.isNull(width) || v > width) {
                    colMaxWidthMap.put(k, v);
                }
            });
        });
        table.setColMaxWidthMap(colMaxWidthMap);
    }

    /**
     * 获取tr中的td
     *
     * @param tr tr
     */
    private void parseTdOfTr(Tr tr) {
        Elements tdElements = tr.getElement().children();
        if (tdElements.isEmpty()) {
            return;
        }
        tr.setTdList(new ArrayList<>());
        tr.setColWidthMap(new HashMap<>(tdElements.size()));
        for (int i = 0, size = tdElements.size(); i < size; i++) {
            Element tdElement = tdElements.get(i);
            Td td = new Td();
            td.setElement(tdElement);
            td.setContent(tdElement.text());
            td.setTh(Objects.equals(TableTag.th.name(), tdElement.tagName()));
            td.setRow(tr.getIndex());
            td.setStyle(StyleUtils.mixStyle(tr.getStyle(), StyleUtils.parseStyle(tdElement)));
            // 除每行第一个单元格外，修正含跨列的单元格位置
            if (i > 0) {
                int shift = tr.getTdList().stream().filter(t -> t.getColSpan() > 0)
                        .mapToInt(t -> t.getColSpan() - 1).sum();
                td.setCol(i + shift);
            } else {
                td.setCol(i);
            }

            String colSpan = tdElement.attr(TableTag.colspan.name());
            td.setColSpan(TdUtils.getSpan(colSpan));

            String rowSpan = tdElement.attr(TableTag.rowspan.name());
            td.setRowSpan(TdUtils.getSpan(rowSpan));

            int rowBound = TdUtils.get(td::getRowSpan, td::getRow);
            td.setRowBound(rowBound);

            int colBound = TdUtils.get(td::getColSpan, td::getCol);
            td.setColBound(colBound);

            tr.getTdList().add(td);

            // 设置每列宽度
            int width = TdUtils.getStringWidth(td.getContent());
            tr.getColWidthMap().put(td.getCol(), width);
        }

        int lastColNumber = tr.getTdList().stream().mapToInt(td -> td.getColSpan() > 0 ? td.getColSpan() : 1).sum();
        tr.setLastColumnNum(lastColNumber);
    }

    /**
     * 调整表格单元格位置
     *
     * @param td      单元格
     * @param trIndex 行索引
     * @param trList  所有行
     */
    private void adjustTdPosition(Td td, int trIndex, List<Tr> trList) {
        List<Td> rowSpanTds = trList.subList(0, trIndex).stream()
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
        int colBound = TdUtils.get(td::getColSpan, td::getCol);
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

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
import java.util.Comparator;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.function.Predicate;
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
        log.info("Start parsing html file");
        long startTime = System.currentTimeMillis();
        parser.document = Jsoup.parse(htmlFile, CharEncoding.UTF_8);
        log.info("Complete html file parsing,takes {} milliseconds", System.currentTimeMillis() - startTime);
        return parser;
    }

    /**
     * 获取所有表格
     *
     * @return 所有表格
     */
    public List<Table> getAllTable() {
        Elements tableElements = document.getElementsByTag(TableTag.table.name());
        return IntStream.range(0, tableElements.size()).mapToObj(i -> {
            Element tableElement = tableElements.get(i);
            Table table = new Table();
            table.setIndex(i);
            table.setElement(tableElement);

            Elements captionElements = tableElement.getElementsByTag(TableTag.caption.name());
            if (!captionElements.isEmpty()) {
                table.setCaption(captionElements.first().text());
            }
            table.setStyleMap(StyleUtils.parseStyle(tableElement));

            this.parseTrOfTable(tableElement, table);
            return table;
        }).collect(Collectors.toList());
    }

    /**
     * 解析table中的tr
     *
     * @param tableElement table元素
     * @param table        table
     */
    private void parseTrOfTable(Element tableElement, Table table) {
        Elements trElements = tableElement.getElementsByTag(TableTag.tr.name());
        if (trElements.isEmpty()) {
            return;
        }
        Map<Element, Map<String, String>> parentStyleMap = new HashMap<>();

        List<Tr> trList = IntStream.range(0, trElements.size()).mapToObj(index -> {
            Element trElement = trElements.get(index);
            Element parent = trElement.parent();
            Map<String, String> upperStyle;
            if (parentStyleMap.containsKey(parent)) {
                upperStyle = parentStyleMap.get(parent);
            } else {
                upperStyle = StyleUtils.mixStyle(table.getStyleMap(), StyleUtils.parseStyle(parent));
                parentStyleMap.put(parent, upperStyle);
            }
            Tr tr = new Tr(index);
            tr.setElement(trElement);
            tr.setStyle(StyleUtils.mixStyle(upperStyle, StyleUtils.parseStyle(trElement)));
            this.parseTdOfTr(trElement, tr);
            return tr;
        }).collect(Collectors.toList());

        table.setTrList(trList);

        int lastColumnNum = trList.parallelStream().max(Comparator.comparing(Tr::getLastColumnNum)).get().getLastColumnNum();
        table.setLastColumnNum(lastColumnNum);

        Map<Integer, Integer> colMaxWidthMap = new HashMap<>();
        trList.stream().map(Tr::getColWidthMap).forEach(map -> {
            map.forEach((k, v) -> {
                Integer width = colMaxWidthMap.get(k);
                if (Objects.isNull(width) || v > width) {
                    colMaxWidthMap.put(k, v);
                }
            });
        });
        table.setColMaxWidthMap(colMaxWidthMap);

        // 调整td位置,排除第一行，第一行不需要进行调整
        if (trList.size() == 1) {
            return;
        }
        trList.subList(1, trList.size()).parallelStream().forEach(tr -> {
            tr.getTdList().parallelStream().forEach(td -> this.adjustTdPosition(td, tr.getIndex(), trList, lastColumnNum));
        });
    }

    /**
     * 获取tr中的td
     *
     * @param trElement tr元素
     * @param tr        tr
     */
    private void parseTdOfTr(Element trElement, Tr tr) {
        Elements tdElements = trElement.children();
        if (tdElements.isEmpty()) {
            return;
        }
        tr.setTdList(new ArrayList<>());
        tr.setColWidthMap(new HashMap<>());
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

            tr.getTdList().add(td);

            // 设置每列宽度
            int width = TdUtils.getStringWidth(td.getContent());
            tr.getColWidthMap().put(td.getCol(), width);

            int colIndex = TdUtils.get(td::getColSpan, td::getCol);
            if (Objects.isNull(tr.getLastColumnNum()) || colIndex > tr.getLastColumnNum()) {
                tr.setLastColumnNum(colIndex);
            }
        }
    }

    /**
     * 调整表格单元格位置
     *
     * @param td            单元格
     * @param trIndex       行索引
     * @param trList        所有行
     * @param lastColumnNum 最后列编号
     */
    private void adjustTdPosition(Td td, int trIndex, List<Tr> trList, int lastColumnNum) {
        Predicate<Tr> predicate = tr -> tr.getIndex() < trIndex;
        List<Td> tds = trList.stream().filter(predicate).flatMap(tr -> tr.getTdList().stream())
                .collect(Collectors.toList());
        if (CollectionUtils.isEmpty(tds)) {
            return;
        }
        for (int i = 0; i < lastColumnNum; i++) {
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

package com.github.liaochong.html2excel.core;

import java.io.File;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.List;
import java.util.Optional;
import java.util.function.Predicate;
import java.util.stream.Collectors;

import org.apache.commons.collections4.CollectionUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;

import com.github.liaochong.html2excel.exception.NoTablesException;
import com.github.liaochong.html2excel.utils.TdUtils;

/**
 * @author liaochong
 * @version 1.0
 */
public class Html2Excel {

    private Document document;

    private Workbook workbook;

    private final List<Tr> TRS = new ArrayList<>();

    private int maxCols;

    private Html2Excel() {
    }

    /**
     * 读取html
     * 
     * @param htmlFile html文件
     * @throws Exception 解析异常
     */
    public static Html2Excel readHtml(File htmlFile) throws Exception {
        Html2Excel html2Excel = new Html2Excel();
        html2Excel.document = Jsoup.parse(htmlFile, StandardCharsets.UTF_8.displayName());
        return html2Excel;
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
        if (CollectionUtils.isNotEmpty(TRS)) {
            TRS.clear();
        }
        Elements trs = table.getElementsByTag(Tag.tr.name());
        for (int i = 0; i < trs.size(); i++) {
            Tr trContainer = new Tr(i);
            TRS.add(trContainer);
            this.processTr(trs.get(i), trContainer);
        }
        Sheet sheet = this.getSheet(table, index);

        List<Td> allTds = this.adjust();
        allTds.forEach(td -> {
            if (td.getRowSpan() == 0 && td.getColSpan() == 0) {
                return;
            }
            sheet.addMergedRegion(new CellRangeAddress(td.getX(), TdUtils.get(td::getRowSpan, td::getX), td.getY(),
                    TdUtils.get(td::getColSpan, td::getY)));
        });
        allTds.forEach(td -> {
            Row row = sheet.getRow(td.getX());
            Cell cell = row.getCell(td.getY());
            cell.setCellValue(td.getContent());
        });
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
        for (int i = 0; i < TRS.size(); i++) {
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

            if (i > maxCols) {
                maxCols = i;
            }

            Element element = elements.get(i);
            String colSpan = element.attr(Tag.colspan.name());
            if (StringUtils.isNotBlank(colSpan)) {
                td.setColSpan(Integer.valueOf(colSpan));
            }
            String rowSpan = element.attr(Tag.rowspan.name());
            if (StringUtils.isNotBlank(rowSpan)) {
                td.setRowSpan(Integer.valueOf(rowSpan));
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
        List<Td> allTds = TRS.stream().flatMap(tr -> tr.getTds().stream()).collect(Collectors.toList());
        TRS.forEach(tr -> tr.getTds().parallelStream().forEach(td -> this.adjust(allTds, td)));
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

package com.github.liaochong.html2excel.core;

import java.io.File;
import java.io.FileOutputStream;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.List;
import java.util.Optional;
import java.util.function.Predicate;
import java.util.stream.Collectors;

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

/**
 * @author liaochong
 * @version 1.0
 */
public class Html2Excel {

    private static final List<Tr> TRS = new ArrayList<>();

    private int maxCols;

    public static void main(String[] args) throws Exception {
        Html2Excel html2Excel = new Html2Excel();
        html2Excel.readDocument(new File("/Users/liaochong/Downloads/1.html"));
    }

    public void readDocument(File htmlFile) throws Exception {
        Document document = Jsoup.parse(htmlFile, StandardCharsets.UTF_8.displayName());

        Elements tables = document.getElementsByTag("table");
        tables.forEach(this::processTable);
    }

    private void processTable(Element table) {
        Elements trs = table.getElementsByTag("tr");
        for (int i = 0; i < trs.size(); i++) {
            Tr trContainer = new Tr(i);
            TRS.add(trContainer);
            this.processTr(trs.get(i), trContainer);
        }
        List<Td> allTds = TRS.stream().flatMap(tr -> tr.getTds().stream()).collect(Collectors.toList());
        for (int i = 1; i < TRS.size(); i++) {
            List<Td> tds = TRS.get(i).getTds();
            tds.forEach(td -> {
                Predicate<Td> predicate = prevTd -> prevTd.getX() < td.getX() && prevTd.getY() == td.getY()
                        && (prevTd.getRowSpan() > 0 ? prevTd.getX() + prevTd.getRowSpan() - 1 : prevTd.getX()) >= td
                                .getX();
                Optional<Td> findResult = allTds.stream().filter(predicate).findAny();
                if (findResult.isPresent()) {
                    Td sameColTd = findResult.get();
                    int prevTdColSpan = sameColTd.getColSpan();
                    int realY = prevTdColSpan > 0 ? td.getY() + prevTdColSpan : td.getY() + 1;
                    td.setY(realY);
                }
            });
        }

        Workbook wb = new XSSFWorkbook();
        Sheet sheet = wb.createSheet("1");
        for (int i = 0; i < TRS.size(); i++) {
            Row row = sheet.createRow(i);
            for (int j = 0; j <= maxCols; j++) {
                row.createCell(j);
            }
        }
        allTds.forEach(td -> {
            if (td.getRowSpan() > 0 || td.getColSpan() > 0) {
                sheet.addMergedRegion(new CellRangeAddress(td.getX(),
                        td.getRowSpan() > 0 ? td.getX() + td.getRowSpan() - 1 : td.getX(), td.getY(),
                        td.getColSpan() > 0 ? td.getY() + td.getColSpan() - 1 : td.getY()));
            }
        });
        allTds.forEach(td -> {
            Row row = sheet.getRow(td.getX());
            Cell cell = row.getCell(td.getY());
            cell.setCellValue(td.getContent());
        });
        try (FileOutputStream fileOut = new FileOutputStream("/Users/liaochong/Develop/workbook.xlsx");) {
            wb.write(fileOut);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private void processTr(Element tr, Tr container) {
        Elements ths = tr.getElementsByTag("th");
        if (!ths.isEmpty()) {
            this.processing(ths, container, true);
        }
        Elements tds = tr.getElementsByTag("td");
        if (!tds.isEmpty()) {
            this.processing(tds, container, false);
        }
    }

    private void processing(Elements elements, Tr container, boolean isTh) {
        for (int i = 0; i < elements.size(); i++) {
            Td td = new Td();
            td.setTh(isTh);
            td.setX(container.getIndex());
            td.setY(i);

            if (i > maxCols) {
                maxCols = i;
            }

            Element element = elements.get(i);
            String colSpan = element.attr("colspan");
            String rowSpan = element.attr("rowspan");
            if (StringUtils.isNotBlank(colSpan)) {
                td.setColSpan(Integer.valueOf(colSpan));
            }
            if (StringUtils.isNotBlank(rowSpan)) {
                td.setRowSpan(Integer.valueOf(rowSpan));
            }
            td.setContent(element.text());
            container.getTds().add(td);
        }
    }
}

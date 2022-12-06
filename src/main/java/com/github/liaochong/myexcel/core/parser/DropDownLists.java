package com.github.liaochong.myexcel.core.parser;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.Map;
import java.util.Objects;
import java.util.concurrent.ConcurrentHashMap;

public final class DropDownLists {

    private static final Map<String, Index> INDEX = new ConcurrentHashMap<>();

    private DropDownLists() {
    }

    public static Index getHiddenSheetIndex(String input, Workbook workbook) {
        return INDEX.computeIfAbsent(input, hash -> createAndWriteHiddenSheet(input, workbook));
    }

    private static Index createAndWriteHiddenSheet(String input, Workbook workbook) {
        String sheetName = "MyExcel_HiddenDat@List-0";
        Sheet sheet = workbook.getSheet(sheetName);
        if (sheet == null) {
            sheet = workbook.createSheet(sheetName);
            int index = workbook.getSheetIndex(sheet);
            workbook.setSheetHidden(index, true);
        }
        int rowNum = sheet.getLastRowNum() + 1;
        Row row = sheet.createRow(rowNum);
        String[] list = input.split(",");
        for (int i = 0; i < list.length; i++) {
            Cell cell = row.createCell(i);
            cell.setCellValue(list[i]);
        }
        int displayRowNum = rowNum + 1;
        return new Index(list[0], rowNum, "'" + sheetName + "'!$" + displayRowNum + ":$" + displayRowNum);
    }

    public static class Index {

        public String firstLine;

        public int rowNum;

        public String path;

        public Index(String firstLine, int rowNum, String path) {
            this.firstLine = firstLine;
            this.rowNum = rowNum;
            this.path = path;
        }

        @Override
        public boolean equals(Object o) {
            if (this == o) return true;
            if (!(o instanceof Index)) return false;
            Index index = (Index) o;
            return rowNum == index.rowNum && Objects.equals(firstLine, index.firstLine) && Objects.equals(path, index.path);
        }

        @Override
        public int hashCode() {
            return Objects.hash(firstLine, rowNum, path);
        }
    }
}

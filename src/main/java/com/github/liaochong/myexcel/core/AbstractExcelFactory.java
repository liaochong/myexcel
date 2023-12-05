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
package com.github.liaochong.myexcel.core;

import com.github.liaochong.myexcel.core.parser.ContentTypeEnum;
import com.github.liaochong.myexcel.core.parser.DropDownLists;
import com.github.liaochong.myexcel.core.parser.HtmlTableParser;
import com.github.liaochong.myexcel.core.parser.Td;
import com.github.liaochong.myexcel.core.parser.Tr;
import com.github.liaochong.myexcel.core.strategy.SheetStrategy;
import com.github.liaochong.myexcel.core.strategy.WidthStrategy;
import com.github.liaochong.myexcel.core.style.BackgroundStyle;
import com.github.liaochong.myexcel.core.style.BorderStyle;
import com.github.liaochong.myexcel.core.style.CustomColor;
import com.github.liaochong.myexcel.core.style.FontStyle;
import com.github.liaochong.myexcel.core.style.LinkDefaultCellStyle;
import com.github.liaochong.myexcel.core.style.TdDefaultCellStyle;
import com.github.liaochong.myexcel.core.style.TextAlignStyle;
import com.github.liaochong.myexcel.core.style.ThDefaultCellStyle;
import com.github.liaochong.myexcel.core.style.WordBreakStyle;
import com.github.liaochong.myexcel.exception.ExcelBuildException;
import com.github.liaochong.myexcel.exception.SaxReadException;
import com.github.liaochong.myexcel.utils.ColorUtil;
import com.github.liaochong.myexcel.utils.StringUtil;
import com.github.liaochong.myexcel.utils.TdUtil;
import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.hssf.usermodel.HSSFPalette;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.FileMagic;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.ss.usermodel.DataValidationConstraint;
import org.apache.poi.ss.usermodel.DataValidationHelper;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Hyperlink;
import org.apache.poi.ss.usermodel.Picture;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.ShapeTypes;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFDataValidation;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFSimpleShape;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.util.Collections;
import java.util.EnumMap;
import java.util.HashMap;
import java.util.LinkedList;
import java.util.List;
import java.util.Map;
import java.util.Objects;

/**
 * @author liaochong
 * @version 1.0
 */
public abstract class AbstractExcelFactory implements ExcelFactory {

    private static final Logger logger = LoggerFactory.getLogger(AbstractExcelFactory.class);

    protected static final int XLSX_MAX_ROW_COUNT = 1048576;

    protected static final int XLS_MAX_ROW_COUNT = 65536;

    private static Map<String, String> DEFAULT_TD_STYLE;

    private static Map<String, String> DEFAULT_TH_STYLE;

    private static Map<String, String> DEFAULT_LINK_STYLE;

    protected Workbook workbook;
    /**
     * 是否为hssf
     */
    protected boolean isHssf;
    /**
     * 每行的单元格最大高度map
     */
    private Map<Integer, Short> maxTdHeightMap = new HashMap<>();
    /**
     * 是否使用默认样式
     */
    private boolean useDefaultStyle;
    /**
     * 是否应用默认样式，允许覆盖
     */
    private boolean applyDefaultStyle;
    /**
     * 自定义颜色
     */
    private CustomColor customColor;
    /**
     * 单元格样式映射
     */
    private Map<Map<String, String>, CellStyle> cellStyleMap = new HashMap<>();
    /**
     * 样式容器
     */
    private Map<HtmlTableParser.HtmlTag, CellStyle> defaultCellStyleMap;
    /**
     * 字体map
     */
    private Map<String, Font> fontMap = new HashMap<>();
    /**
     * 冻结区域
     */
    private FreezePane[] freezePanes;
    /**
     * 自动宽度策略，默认无宽度策略
     */
    protected WidthStrategy widthStrategy = WidthStrategy.NO_AUTO;
    /**
     * 生成sheet策略，默认生成多个sheet
     */
    protected SheetStrategy sheetStrategy = SheetStrategy.MULTI_SHEET;
    /**
     * 暂存单元格，由后续行认领
     */
    protected List<Td> stagingTds = new LinkedList<>();

    private CreationHelper createHelper;

    private DataFormat format;
    /**
     * 图片路径缓存
     */
    private Map<String, Integer> imageMapping;

    @Override
    public ExcelFactory useDefaultStyle() {
        this.useDefaultStyle = true;
        return this;
    }

    @Override
    public ExcelFactory applyDefaultStyle() {
        this.applyDefaultStyle = true;
        if (DEFAULT_TD_STYLE == null) {
            DEFAULT_TD_STYLE = new HashMap<String, String>() {{
                put("text-align", "center");
                put("vertical-align", "center");
                put("border-style", "thin");
            }};
            DEFAULT_TH_STYLE = new HashMap<String, String>(DEFAULT_TD_STYLE) {{
                put("font-weight", "bold");
            }};
            DEFAULT_LINK_STYLE = new HashMap<String, String>(DEFAULT_TD_STYLE) {{
                put("text-decoration", "underline");
                put("color", "blue");
            }};
        }
        return this;
    }

    @Override
    public ExcelFactory freezePanes(FreezePane... freezePanes) {
        this.freezePanes = freezePanes;
        return this;
    }

    @Override
    public ExcelFactory workbookType(WorkbookType workbookType) {
        if (Objects.nonNull(workbook)) {
            return this;
        }
        switch (workbookType) {
            case XLS:
                workbook = new HSSFWorkbook();
                isHssf = true;
                break;
            case SXLSX:
                workbook = new SXSSFWorkbook(1);
                break;
            default:
                workbook = new XSSFWorkbook();
        }
        return this;
    }

    @Override
    public ExcelFactory widthStrategy(WidthStrategy widthStrategy) {
        this.widthStrategy = widthStrategy;
        return this;
    }

    @Override
    public ExcelFactory sheetStrategy(SheetStrategy sheetStrategy) {
        this.sheetStrategy = sheetStrategy;
        return this;
    }

    protected String getRealSheetName(String sheetName) {
        if (sheetName == null) {
            sheetName = "Sheet";
        }
        Sheet sheet = workbook.getSheet(sheetName);
        int sort = 1;
        String realSheetName = sheetName;
        while (sheet != null) {
            realSheetName = sheetName + " (" + sort + ")";
            sheet = workbook.getSheet(realSheetName);
            sort++;
        }
        return realSheetName;
    }

    /**
     * 创建行-row
     *
     * @param tr    tr
     * @param sheet sheet
     */
    protected void createRow(Tr tr, Sheet sheet) {
        Row row = sheet.createRow(tr.index);
        if (!tr.visibility) {
            row.setZeroHeight(true);
        }
        if (tr.height > 0) {
            row.setHeightInPoints(tr.height);
        } else {
            // 设置行高，最小12
            if (maxTdHeightMap.get(row.getRowNum()) == null) {
                row.setHeightInPoints(row.getHeightInPoints() + 5);
            } else {
                row.setHeightInPoints((short) (maxTdHeightMap.get(row.getRowNum()) + 5));
                maxTdHeightMap.remove(row.getRowNum());
            }
        }
        stagingTds.stream().filter(blankTd -> Objects.equals(blankTd.row, tr.index)).forEach(td -> {
            if (tr.tdList == Collections.EMPTY_LIST) {
                tr.tdList = new LinkedList<>();
            }
            tr.tdList.add(td);
        });
        for (Td td : tr.tdList) {
            this.createCell(td, sheet, row);
            if (td.rowSpan == 0) {
                continue;
            }
            for (int i = td.row + 1, rowBound = td.getRowBound(); i <= rowBound; i++) {
                for (int j = td.col, colBound = td.getColBound(); j <= colBound; j++) {
                    Td blankTd = new Td(i, j);
                    blankTd.th = td.th;
                    blankTd.style = td.style;
                    stagingTds.add(blankTd);
                }
            }
        }
        // 移除暂存区空白单元格
        stagingTds.removeIf(td -> Objects.equals(td.row, tr.index));
    }

    /**
     * 创建单元格
     *
     * @param td         td
     * @param sheet      sheet
     * @param currentRow 当前行
     */
    protected void createCell(Td td, Sheet sheet, Row currentRow) {
        Cell cell;
        if (td.formula) {
            cell = currentRow.createCell(td.col, CellType.FORMULA);
            cell.setCellFormula(td.content);
        } else {
            String content = td.content;
            switch (td.tdContentType) {
                case DOUBLE:
                    cell = currentRow.createCell(td.col, CellType.NUMERIC);
                    if (null != content) {
                        cell.setCellValue(Double.parseDouble(content));
                    }
                    break;
                case DATE:
                    cell = currentRow.createCell(td.col);
                    if (td.date != null) {
                        cell.setCellValue(td.date);
                    } else if (td.localDateTime != null) {
                        cell.setCellValue(td.localDateTime);
                    } else if (td.localDate != null) {
                        cell.setCellValue(td.localDate);
                    }
                    break;
                case BOOLEAN:
                    cell = currentRow.createCell(td.col, CellType.BOOLEAN);
                    if (null != content) {
                        cell.setCellValue(Boolean.parseBoolean(content));
                    }
                    break;
                case NUMBER_DROP_DOWN_LIST:
                    cell = currentRow.createCell(td.col, CellType.NUMERIC);
                    String firstEle = setDropDownList(td, sheet, content);
                    if (firstEle != null) {
                        cell.setCellValue(Double.parseDouble(firstEle));
                    }
                    break;
                case BOOLEAN_DROP_DOWN_LIST:
                    cell = currentRow.createCell(td.col, CellType.BOOLEAN);
                    firstEle = setDropDownList(td, sheet, content);
                    if (firstEle != null) {
                        cell.setCellValue(Boolean.parseBoolean(firstEle));
                    }
                    break;
                case DROP_DOWN_LIST:
                    cell = currentRow.createCell(td.col, CellType.STRING);
                    firstEle = setDropDownList(td, sheet, content);
                    if (firstEle != null) {
                        cell.setCellValue(firstEle);
                    }
                    break;
                case LINK_URL:
                    cell = setLink(td, currentRow, HyperlinkType.URL);
                    break;
                case LINK_EMAIL:
                    cell = setLink(td, currentRow, HyperlinkType.EMAIL);
                    break;
                case IMAGE:
                    cell = currentRow.createCell(td.col);
                    setImage(td, sheet);
                    break;
                default:
                    cell = currentRow.createCell(td.col, CellType.STRING);
                    cell.setCellValue(content);
                    break;
            }
            this.setPrompt(td, sheet);
        }
        // 设置斜线
        this.drawingSlant(td, sheet);
        // 设置批注
        this.setComment(td, sheet, cell);
        // 设置单元格样式
        this.setCellStyle(currentRow, cell, td);
        if (td.col != td.getColBound()) {
            for (int j = td.col + 1, colBound = td.getColBound(); j <= colBound; j++) {
                cell = currentRow.createCell(j);
                this.setCellStyle(currentRow, cell, td);
            }
        }
        if (td.colSpan > 0 || td.rowSpan > 0) {
            sheet.addMergedRegion(new CellRangeAddress(td.row, td.getRowBound(), td.col, td.getColBound()));
        }
    }

    private void setComment(Td td, Sheet sheet, Cell cell) {
        if (td.comment == null) {
            return;
        }
        if (createHelper == null) {
            createHelper = workbook.getCreationHelper();
        }
        Drawing<?> drawing = sheet.getDrawingPatriarch();
        if (drawing == null) {
            drawing = sheet.createDrawingPatriarch();
        }
        ClientAnchor anchor = createHelper.createClientAnchor();
        anchor.setCol1(cell.getColumnIndex());
        anchor.setCol2(cell.getColumnIndex() + 2);
        anchor.setRow1(td.row);
        anchor.setRow2(td.getRowBound() + 2);
        Comment comment = drawing.createCellComment(anchor);
        RichTextString str = createHelper.createRichTextString(td.comment.text);
        comment.setString(str);
        comment.setAuthor(td.comment.author);
        cell.setCellComment(comment);
    }

    private void drawingSlant(Td td, Sheet sheet) {
        if (td.slant == null) {
            return;
        }
        if (isHssf) {
            throw new UnsupportedOperationException("The current workbook does not support setting slashes.");
        }
        XSSFDrawing drawing;
        if (workbook instanceof SXSSFWorkbook) {
            drawing = ((SXSSFSheet) sheet).getDrawingPatriarch();
            if (drawing == null) {
                sheet.createDrawingPatriarch();
                drawing = ((SXSSFSheet) sheet).getDrawingPatriarch();
            }
        } else {
            drawing = ((XSSFSheet) sheet).getDrawingPatriarch();
            if (drawing == null) {
                drawing = ((XSSFSheet) sheet).createDrawingPatriarch();
            }
        }
        if (createHelper == null) {
            createHelper = workbook.getCreationHelper();
        }
        ClientAnchor anchor = createHelper.createClientAnchor();
        // 设置斜线的开始位置
        anchor.setCol1(td.col);
        anchor.setRow1(td.row);
        // 设置斜线的结束位置
        anchor.setCol2(td.getColBound() + 1);
        anchor.setRow2(td.getRowBound() + 1);
        XSSFSimpleShape shape = drawing.createSimpleShape((XSSFClientAnchor) anchor);
        // 设置形状类型为线型
        shape.setShapeType(ShapeTypes.LINE);
        // 设置线宽
        shape.setLineWidth(td.slant.lineWidth);
        // 设置线的风格
        shape.setLineStyle(td.slant.lineStyle);
        // 设置线的颜色
        int[] color = ColorUtil.getRGBByColor(td.slant.lineStyleColor);
        shape.setLineStyleColor(color[0], color[1], color[2]);
    }

    private void setPrompt(Td td, Sheet sheet) {
        if (td.promptContainer == null) {
            return;
        }
        if (ContentTypeEnum.isDropdownList(td.tdContentType)) {
            return;
        }
        DataValidationHelper dvHelper = sheet.getDataValidationHelper();
        DataValidationConstraint constraint = dvHelper.createCustomConstraint("BB1");
        CellRangeAddressList addressList = new CellRangeAddressList(
                td.row, td.getRowBound(), td.col, td.getColBound());
        DataValidation dataValidation = dvHelper.createValidation(constraint, addressList);
        dataValidation.createPromptBox(td.promptContainer.title, td.promptContainer.text);
        dataValidation.setShowPromptBox(true);
        sheet.addValidationData(dataValidation);
    }

    private void setImage(Td td, Sheet sheet) {
        if (td.file == null && td.fileIs == null) {
            return;
        }
        if (createHelper == null) {
            createHelper = workbook.getCreationHelper();
        }
        int pictureIdx;
        int format;
        if (td.file != null) {
            String fileName = td.file.getName();
            String suffix = fileName.substring(fileName.lastIndexOf(".") + 1).toLowerCase();
            switch (suffix) {
                case "jpg":
                case "jpeg":
                    format = Workbook.PICTURE_TYPE_JPEG;
                    break;
                case "png":
                    format = Workbook.PICTURE_TYPE_PNG;
                    break;
                case "dib":
                    format = Workbook.PICTURE_TYPE_DIB;
                    break;
                case "emf":
                    format = Workbook.PICTURE_TYPE_EMF;
                    break;
                case "pict":
                    format = Workbook.PICTURE_TYPE_PICT;
                    break;
                case "wmf":
                    format = Workbook.PICTURE_TYPE_WMF;
                    break;
                default:
                    throw new IllegalArgumentException("Invalid image type");
            }
            if (imageMapping == null) {
                imageMapping = new HashMap<>();
            }
            pictureIdx = imageMapping.computeIfAbsent(td.file.getAbsolutePath(), s -> {
                try {
                    byte[] bytes = Files.readAllBytes(td.file.toPath());
                    return workbook.addPicture(bytes, format);
                } catch (IOException e) {
                    logger.error("read image failure", e);
                    throw new ExcelBuildException("read image failure, path:" + td.file.getAbsolutePath(), e);
                }
            });
        } else {
            FileMagic fm;
            try (InputStream is = FileMagic.prepareToCheckMagic(td.fileIs)) {
                fm = FileMagic.valueOf(is);
                switch (fm) {
                    case JPEG:
                        format = Workbook.PICTURE_TYPE_JPEG;
                        break;
                    case PNG:
                        format = Workbook.PICTURE_TYPE_PNG;
                        break;
                    case EMF:
                        format = Workbook.PICTURE_TYPE_EMF;
                        break;
                    case WMF:
                        format = Workbook.PICTURE_TYPE_WMF;
                        break;
                    default:
                        throw new IllegalArgumentException("Invalid image type");
                }
                pictureIdx = workbook.addPicture(IOUtils.toByteArray(is), format);
            } catch (Throwable throwable) {
                throw new SaxReadException("Fail to get excel magic", throwable);
            }
        }
        ClientAnchor anchor = createHelper.createClientAnchor();
        anchor.setAnchorType(ClientAnchor.AnchorType.MOVE_AND_RESIZE);
        anchor.setDx1(0);
        anchor.setDy1(0);
        final int emuPerMm = 36000;
        anchor.setDx2(isHssf ? 1023 : 100 * emuPerMm);
        anchor.setDy2(isHssf ? 1023 : 99 * emuPerMm);
        anchor.setCol1(td.col);
        anchor.setRow1(td.row);
        anchor.setCol2(td.getColBound());
        anchor.setRow2(td.getRowBound());
        Drawing<?> drawing = sheet.getDrawingPatriarch();
        if (drawing == null) {
            drawing = sheet.createDrawingPatriarch();
        }
        Picture pict = drawing.createPicture(anchor, pictureIdx);
        // only support JPEG and PNG
        if (td.getImage() != null) {
            pict.resize(td.getImage().getScaleX(), td.getImage().getScaleY());
        }
    }

    private Cell setLink(Td td, Row currentRow, HyperlinkType hyperlinkType) {
        if (StringUtil.isBlank(td.content)) {
            return currentRow.createCell(td.col);
        }
        if (createHelper == null) {
            createHelper = workbook.getCreationHelper();
        }
        Cell cell = currentRow.createCell(td.col, CellType.STRING);
        cell.setCellValue(td.content);
        Hyperlink link = createHelper.createHyperlink(hyperlinkType);
        link.setAddress(td.link);
        cell.setHyperlink(link);
        return cell;
    }

    private String setDropDownList(Td td, Sheet sheet, String content) {
        if (content != null && !content.isEmpty()) {
            CellRangeAddressList addressList = new CellRangeAddressList(
                    td.row, td.getRowBound(), td.col, td.getColBound());
            DataValidationHelper dvHelper = sheet.getDataValidationHelper();
            String[] list;
            DataValidation validation;
            if (content.length() <= 250) {
                list = content.split(",");
                DataValidationConstraint dvConstraint = dvHelper.createExplicitListConstraint(list);
                validation = dvHelper.createValidation(
                        dvConstraint, addressList);

            } else {
                DropDownLists.Index index = DropDownLists.getHiddenSheetIndex(content, workbook);
                list = new String[]{index.firstLine};
                validation = dvHelper.createValidation(dvHelper.createFormulaListConstraint(index.path), addressList);
            }

            if (td.promptContainer != null) {
                validation.createPromptBox(td.promptContainer.title, td.promptContainer.text);
                validation.setShowPromptBox(true);
            }
            if (validation instanceof XSSFDataValidation) {
                validation.setSuppressDropDownArrow(true);
                validation.setShowErrorBox(true);
            } else {
                validation.setSuppressDropDownArrow(false);
            }
            sheet.addValidationData(validation);
            if (list.length > 0) {
                return list[0];
            }
        }
        return null;
    }

    /**
     * 设置单元格样式
     *
     * @param cell 单元格
     * @param td   td单元格
     */
    private void setCellStyle(Row row, Cell cell, Td td) {
        if (useDefaultStyle) {
            if (td.th) {
                cell.setCellStyle(defaultCellStyleMap.get(HtmlTableParser.HtmlTag.th));
            } else {
                if (ContentTypeEnum.isLink(td.tdContentType)) {
                    cell.setCellStyle(defaultCellStyleMap.get(HtmlTableParser.HtmlTag.link));
                } else {
                    cell.setCellStyle(defaultCellStyleMap.get(HtmlTableParser.HtmlTag.td));
                }
            }
        } else {
            this.doSetInnerSpan(cell, td);
            if (td.style.isEmpty() && !applyDefaultStyle) {
                return;
            }
            String fs = td.style.get("font-size");
            if (fs != null) {
                short fontSize = (short) TdUtil.getValue(fs);
                if (fontSize > maxTdHeightMap.getOrDefault(row.getRowNum(), FontStyle.DEFAULT_FONT_SIZE)) {
                    maxTdHeightMap.put(row.getRowNum(), fontSize);
                }
            }
            if (applyDefaultStyle) {
                if (td.th) {
                    DEFAULT_TH_STYLE.forEach((k, v) -> td.style.putIfAbsent(k, v));
                } else {
                    if (ContentTypeEnum.isLink(td.tdContentType)) {
                        DEFAULT_LINK_STYLE.forEach((k, v) -> td.style.putIfAbsent(k, v));
                    } else {
                        DEFAULT_TD_STYLE.forEach((k, v) -> td.style.putIfAbsent(k, v));
                    }
                }
            }
            if (cellStyleMap.containsKey(td.style)) {
                cell.setCellStyle(cellStyleMap.get(td.style));
                return;
            }
            CellStyle cellStyle = workbook.createCellStyle();
            // background-color
            BackgroundStyle.setBackgroundColor(cellStyle, td.style, customColor);
            // text-align
            TextAlignStyle.setTextAlign(cellStyle, td.style);
            // border
            BorderStyle.setBorder(cellStyle, td.style);
            // word-break
            WordBreakStyle.setWordBreak(cellStyle, td.style);
            // 内容格式
            String formatStr = td.style.get("format");
            if (formatStr != null) {
                if (format == null) {
                    format = workbook.createDataFormat();
                }
                cellStyle.setDataFormat(format.getFormat(formatStr));
            }
            // font
            if (td.fonts == null || td.fonts.isEmpty()) {
                FontStyle.setFont(() -> workbook.createFont(), cellStyle, td.style, fontMap, customColor);
            }
            cell.setCellStyle(cellStyle);
            cellStyleMap.put(td.style, cellStyle);
        }
    }

    private void doSetInnerSpan(Cell cell, Td td) {
        if (td.fonts == null || td.fonts.isEmpty()) {
            return;
        }
        RichTextString richText = isHssf ? new HSSFRichTextString(td.content) : new XSSFRichTextString(td.content);
        for (com.github.liaochong.myexcel.core.parser.Font font : td.fonts) {
            Font f = FontStyle.getFont(font.style, fontMap, () -> workbook.createFont(), customColor);
            richText.applyFont(font.startIndex, font.endIndex, f);
        }
        cell.setCellValue(richText);
    }

    /**
     * 空工作簿
     *
     * @return Workbook
     */
    protected Workbook emptyWorkbook() {
        if (workbook == null) {
            workbook = new XSSFWorkbook();
        }
        workbook.createSheet();
        return workbook;
    }

    /**
     * 初始化默认单元格样式
     *
     * @param workbook workbook
     */
    protected void initCellStyle(Workbook workbook) {
        if (useDefaultStyle) {
            defaultCellStyleMap = new EnumMap<>(HtmlTableParser.HtmlTag.class);
            defaultCellStyleMap.put(HtmlTableParser.HtmlTag.th, new ThDefaultCellStyle().supply(workbook));
            defaultCellStyleMap.put(HtmlTableParser.HtmlTag.td, new TdDefaultCellStyle().supply(workbook));
            defaultCellStyleMap.put(HtmlTableParser.HtmlTag.link, new LinkDefaultCellStyle().supply(workbook));
        } else {
            if (workbook instanceof HSSFWorkbook) {
                HSSFPalette palette = ((HSSFWorkbook) workbook).getCustomPalette();
                customColor = new CustomColor(true, palette);
            } else {
                customColor = new CustomColor();
            }
        }
    }

    /**
     * 窗口冻结
     *
     * @param tableIndex table index
     * @param sheet      sheet
     */
    protected void freezePane(int tableIndex, Sheet sheet) {
        if (freezePanes != null && freezePanes.length > tableIndex) {
            FreezePane freezePane = freezePanes[tableIndex];
            if (freezePane == null) {
                throw new IllegalStateException("FreezePane is null");
            }
            sheet.createFreezePane(freezePane.colSplit, freezePane.rowSplit);
        }
    }

    /**
     * 获取每列最大宽度
     *
     * @param trList trList
     * @return colMaxWidthMap
     */
    protected Map<Integer, Integer> getColMaxWidthMap(List<Tr> trList) {
        if (useDefaultStyle) {
            // 使用默认样式，需要重新修正加粗的标题自适应宽度
            trList.parallelStream().forEach(tr -> {
                tr.tdList.stream().filter(td -> td.th).forEach(th -> {
                    int tdWidth = TdUtil.getStringWidth(th.content, 0.25);
                    tr.colWidthMap.put(th.col, tdWidth);
                });
            });
        }
        int mapMaxSize = trList.stream().mapToInt(tr -> tr.colWidthMap.size()).max().orElse(16);
        Map<Integer, Integer> colMaxWidthMap = new HashMap<>(mapMaxSize);
        trList.forEach(tr -> {
            tr.colWidthMap.forEach((k, v) -> {
                Integer width = colMaxWidthMap.get(k);
                if (Objects.isNull(width) || v > width) {
                    colMaxWidthMap.put(k, v);
                }
            });
            tr.colWidthMap = null;
        });
        return colMaxWidthMap;
    }

    /**
     * 设置每列宽度
     *
     * @param colMaxWidthMap 列最大宽度Map
     * @param sheet          sheet
     * @param maxColIndex    最大列索引
     */
    protected void setColWidth(Map<Integer, Integer> colMaxWidthMap, Sheet sheet, int maxColIndex) {
        if (WidthStrategy.isAutoWidth(widthStrategy)) {
            if (sheet instanceof SXSSFSheet) {
                ((SXSSFSheet) sheet).trackAllColumnsForAutoSizing();
            }
            for (int i = 0; i <= maxColIndex; i++) {
                sheet.autoSizeColumn(i);
            }
            if (sheet instanceof SXSSFSheet) {
                ((SXSSFSheet) sheet).untrackAllColumnsForAutoSizing();
            }
        } else {
            colMaxWidthMap.forEach((key, value) -> {
                int contentLength = value << 1;
                if (contentLength > 255) {
                    contentLength = 255;
                }
                sheet.setColumnWidth(key, contentLength << 8);
            });
        }
    }

    /**
     * 清除缓存
     */
    protected void clearCache() {
        cellStyleMap = new HashMap<>();
        fontMap = new HashMap<>();
        maxTdHeightMap = new HashMap<>();
        format = null;
        createHelper = null;
        imageMapping = null;
    }

    /**
     * 关闭工作簿
     */
    protected void closeWorkbook() {
        if (workbook == null) {
            return;
        }
        try {
            if (workbook instanceof SXSSFWorkbook) {
                ((SXSSFWorkbook) workbook).dispose();
            }
            workbook.close();
        } catch (IOException e1) {
            e1.printStackTrace();
        }
    }
}

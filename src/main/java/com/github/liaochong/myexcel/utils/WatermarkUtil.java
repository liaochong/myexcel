/*
 * Copyright 2019 liaochong
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 * http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
package com.github.liaochong.myexcel.utils;

import com.github.liaochong.myexcel.core.watermark.FontImage;
import com.github.liaochong.myexcel.core.watermark.Watermark;
import com.github.liaochong.myexcel.exception.ExcelBuildException;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ooxml.POIXMLDocumentPart;
import org.apache.poi.openxml4j.opc.PackagePartName;
import org.apache.poi.openxml4j.opc.PackageRelationship;
import org.apache.poi.openxml4j.opc.TargetMode;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFRelation;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.ByteArrayOutputStream;
import java.io.IOException;

/**
 * 水印工具
 *
 * @author liaochong
 * @version 1.0
 */
public class WatermarkUtil {

    public static void addWatermark(Workbook workbook, String text) {
        Watermark watermark = new Watermark();
        watermark.setText(text);
        addWatermark(workbook, watermark);
    }

    /**
     * 添加水印
     *
     * @param workbook  工作簿
     * @param watermark 水印文本
     */
    public static void addWatermark(Workbook workbook, Watermark watermark) {
        if (watermark == null) {
            return;
        }
        try {
            if (workbook instanceof HSSFWorkbook) {
                throw new ExcelBuildException("Watermark can only be provided to XSSFWork.");
            }
            BufferedImage image = FontImage.createWatermarkImage(watermark);
            ByteArrayOutputStream os = new ByteArrayOutputStream();
            try {
                ImageIO.write(image, "png", os);
            } catch (IOException e) {
                throw new RuntimeException(e);
            }
            XSSFWorkbook xssfWorkbook = null;
            try {
                xssfWorkbook = workbook instanceof SXSSFWorkbook ? ((SXSSFWorkbook) workbook).getXSSFWorkbook() : ((XSSFWorkbook) workbook);
                int pictureIdx = workbook.addPicture(os.toByteArray(), Workbook.PICTURE_TYPE_PNG);
                POIXMLDocumentPart poixmlDocumentPart = xssfWorkbook.getAllPictures().get(pictureIdx);
                //获取每个Sheet表
                for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                    XSSFSheet sheet = xssfWorkbook.getSheetAt(i);
                    PackagePartName ppn = poixmlDocumentPart.getPackagePart().getPartName();
                    String relType = XSSFRelation.IMAGES.getRelation();
                    //add relation from sheet to the picture data
                    PackageRelationship pr = sheet.getPackagePart().addRelationship(ppn, TargetMode.INTERNAL, relType, null);
                    //set background picture to sheet
                    sheet.getCTWorksheet().addNewPicture().setId(pr.getId());
                }
            } finally {
                clear(xssfWorkbook);
            }
        } finally {
            clear(workbook);
        }
    }

    private static void clear(Workbook workbook) {
        if (workbook == null) {
            return;
        }
        if (workbook instanceof SXSSFWorkbook) {
            ((SXSSFWorkbook) workbook).dispose();
        }
        try {
            workbook.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}

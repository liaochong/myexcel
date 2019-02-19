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
package com.github.liaochong.html2excel.utils;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;

/**
 * 文件导出工具类
 *
 * @author liaochong
 * @version 1.0
 */
public final class FileExportUtil {

    /**
     * 导出
     *
     * @param workbook workbook
     * @param file     file
     * @throws IOException IOException
     */
    public static void export(Workbook workbook, File file) throws IOException {
        String suffix = ".xlsx";
        if (workbook instanceof HSSFWorkbook) {
            suffix = ".xls";
        }
        if (!file.getName().endsWith(suffix)) {
            file.renameTo(new File(file.getName() + suffix));
        }
        try (OutputStream os = new FileOutputStream(file)) {
            workbook.write(os);
        } finally {
            workbook.close();
        }
    }
}

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

import org.apache.commons.codec.CharEncoding;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.net.URLEncoder;

/**
 * 附件导出工具类
 *
 * @author liaochong
 * @version 1.0
 */
public final class AttachmentExportUtil {

    /**
     * 导出
     *
     * @param workbook workbook
     * @param fileName file name,suffix is not required,and it is not recommended to carry a suffix
     * @param response HttpServletResponse
     * @throws IOException IOException
     */
    public static void export(Workbook workbook, String fileName, HttpServletResponse response) throws IOException {
        String suffix = ".xlsx";
        if (workbook instanceof HSSFWorkbook) {
            if (fileName.endsWith(suffix)) {
                fileName = fileName.substring(0, fileName.length() - 1);
            }
            suffix = ".xls";
        }
        if (!fileName.endsWith(suffix)) {
            fileName += suffix;
        }
        response.setCharacterEncoding(CharEncoding.UTF_8);
        response.addHeader("Content-Disposition", "attachment;filename=" + URLEncoder.encode(fileName, CharEncoding.UTF_8));
        workbook.write(response.getOutputStream());
        workbook.close();
    }
}

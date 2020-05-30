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
package com.github.liaochong.myexcel.utils;

import com.github.liaochong.myexcel.core.constant.Constants;
import org.apache.commons.codec.CharEncoding;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.apache.poi.poifs.crypt.EncryptionInfo;
import org.apache.poi.poifs.crypt.EncryptionMode;
import org.apache.poi.poifs.crypt.Encryptor;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.io.OutputStream;
import java.io.UnsupportedEncodingException;
import java.net.URLEncoder;
import java.nio.file.Files;
import java.nio.file.Path;
import java.security.GeneralSecurityException;

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
     */
    public static void export(Workbook workbook, String fileName, HttpServletResponse response) {
        try {
            String suffix = Constants.XLSX;
            if (workbook instanceof HSSFWorkbook) {
                if (fileName.endsWith(suffix)) {
                    fileName = fileName.substring(0, fileName.length() - 1);
                }
                suffix = Constants.XLS;
                response.setContentType("application/vnd.ms-excel");
            } else {
                response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
            }
            if (!fileName.endsWith(suffix)) {
                fileName += suffix;
            }
            setAttachmentConfig(fileName, response);
            workbook.write(response.getOutputStream());
        } catch (IOException e) {
            throw new RuntimeException(e);
        } finally {
            clear(workbook);
        }
    }

    /**
     * 加密导出
     *
     * @param workbook workbook
     * @param fileName fileName
     * @param response response
     * @param password password
     */
    public static void encryptExport(final Workbook workbook, String fileName, HttpServletResponse response, final String password) {
        if (workbook instanceof HSSFWorkbook) {
            throw new IllegalArgumentException("Document encryption for.xls is not supported");
        }
        Path path = null;
        try {
            String suffix = Constants.XLSX;
            path = TempFileOperator.createTempFile("encrypt_temp", suffix);
            workbook.write(Files.newOutputStream(path));

            final POIFSFileSystem fs = new POIFSFileSystem();
            final EncryptionInfo info = new EncryptionInfo(EncryptionMode.standard);
            final Encryptor enc = info.getEncryptor();
            enc.confirmPassword(password);

            try (OPCPackage opc = OPCPackage.open(path.toFile(), PackageAccess.READ_WRITE);
                 OutputStream os = enc.getDataStream(fs)) {
                opc.save(os);
            }
            if (!fileName.endsWith(suffix)) {
                fileName += suffix;
            }
            response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
            setAttachmentConfig(fileName, response);
            fs.writeFilesystem(response.getOutputStream());
        } catch (IOException | InvalidFormatException | GeneralSecurityException e) {
            throw new RuntimeException(e);
        } finally {
            clear(workbook);
            TempFileOperator.deleteTempFile(path);
        }
    }

    private static void clear(Workbook workbook) {
        if (workbook instanceof SXSSFWorkbook) {
            ((SXSSFWorkbook) workbook).dispose();
        }
        try {
            workbook.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    /**
     * 一般文件导出接口
     *
     * @param path     文件
     * @param fileName 导出后文件名称
     * @param response 响应流
     */
    public static void export(Path path, String fileName, HttpServletResponse response) {
        try {
            response.setContentType("application/octet-stream");
            setAttachmentConfig(fileName, response);
            response.getOutputStream().write(Files.readAllBytes(path));
        } catch (IOException e) {
            throw new RuntimeException(e);
        } finally {
            TempFileOperator.deleteTempFile(path);
        }
    }

    private static void setAttachmentConfig(String fileName, HttpServletResponse response) throws UnsupportedEncodingException {
        response.setCharacterEncoding(CharEncoding.UTF_8);
        response.addHeader("Content-Disposition", "attachment;filename=" + URLEncoder.encode(fileName, CharEncoding.UTF_8).replace("+", "%20"));
    }
}

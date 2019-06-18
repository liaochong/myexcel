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
import lombok.experimental.UtilityClass;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.apache.poi.poifs.crypt.EncryptionInfo;
import org.apache.poi.poifs.crypt.EncryptionMode;
import org.apache.poi.poifs.crypt.Encryptor;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;

/**
 * 文件导出工具类
 *
 * @author liaochong
 * @version 1.0
 */
@UtilityClass
public final class FileExportUtil {

    /**
     * 导出
     *
     * @param workbook workbook
     * @param file     file
     * @throws IOException IOException
     */
    public static void export(Workbook workbook, File file) throws IOException {
        String suffix = Constants.XLSX;
        if (workbook instanceof HSSFWorkbook) {
            if (file.getName().endsWith(suffix)) {
                String absolutePath = file.getAbsolutePath();
                file = Paths.get(absolutePath.substring(0, absolutePath.length() - 1)).toFile();
            }
            suffix = Constants.XLS;
        }
        if (!file.getName().endsWith(suffix)) {
            file = Paths.get(file.getAbsolutePath() + suffix).toFile();
        }
        try (OutputStream os = new FileOutputStream(file)) {
            workbook.write(os);
        } finally {
            if (workbook instanceof SXSSFWorkbook) {
                ((SXSSFWorkbook) workbook).dispose();
            }
            workbook.close();
        }
    }

    /**
     * 加密导出
     *
     * @param workbook workbook
     * @param file     file
     * @param password password
     * @throws Exception Exception
     */
    public static void encryptExport(final Workbook workbook, File file, final String password) throws Exception {
        if (workbook instanceof HSSFWorkbook) {
            throw new IllegalArgumentException("Document encryption for.xls is not supported");
        }
        String suffix = Constants.XLSX;
        if (!file.getName().endsWith(suffix)) {
            file = Paths.get(file.getAbsolutePath() + suffix).toFile();
        }
        try (FileOutputStream fos = new FileOutputStream(file)) {
            workbook.write(fos);
            if (workbook instanceof SXSSFWorkbook) {
                ((SXSSFWorkbook) workbook).dispose();
            }

            final POIFSFileSystem fs = new POIFSFileSystem();
            final EncryptionInfo info = new EncryptionInfo(EncryptionMode.standard);
            final Encryptor enc = info.getEncryptor();
            enc.confirmPassword(password);

            try (OPCPackage opc = OPCPackage.open(file, PackageAccess.READ_WRITE);
                 OutputStream os = enc.getDataStream(fs)) {
                opc.save(os);
            }
            try (FileOutputStream fileOutputStream = new FileOutputStream(file)) {
                fs.writeFilesystem(fileOutputStream);
            }
        } finally {
            workbook.close();
        }
    }

    /**
     * 获取workbook输入流
     *
     * @param workbook workbook
     * @return InputStream
     */
    public static InputStream getInputStream(final Workbook workbook) {
        String suffix = Constants.XLSX;
        if (workbook instanceof HSSFWorkbook) {
            suffix = Constants.XLS;
        }
        Path path = TempFileOperator.createTempFile("tem_outs", suffix);
        try {
            workbook.write(Files.newOutputStream(path));
            if (workbook instanceof SXSSFWorkbook) {
                ((SXSSFWorkbook) workbook).dispose();
            }
            return Files.newInputStream(path);
        } catch (IOException e) {
            throw new RuntimeException(e);
        } finally {
            try {
                workbook.close();
            } catch (IOException e) {
                throw new RuntimeException(e);
            }
            TempFileOperator.deleteTempFile(path);
        }
    }
}

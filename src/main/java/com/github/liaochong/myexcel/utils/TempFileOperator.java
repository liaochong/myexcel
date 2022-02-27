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
import com.github.liaochong.myexcel.exception.ExcelBuildException;
import com.github.liaochong.myexcel.exception.SaxReadException;
import org.apache.poi.util.TempFile;
import org.slf4j.Logger;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import java.util.Objects;

/**
 * 临时文件操作类
 *
 * @author liaochong
 * @version 1.0
 */
public class TempFileOperator {

    private static final Logger log = org.slf4j.LoggerFactory.getLogger(TempFileOperator.class);

    /**
     * 依据前缀名称创建临时文件
     *
     * @param prefix 临时文件前缀
     * @param suffix 临时文件后缀
     * @return Path
     */
    public static Path createTempFile(String prefix, String suffix) {
        try {
            return TempFile.createTempFile("liaochong$myexcel_" + prefix + "_" + Thread.currentThread().getId() + "_", suffix).toPath();
        } catch (IOException e) {
            throw ExcelBuildException.of("Failed to create temp file", e);
        }
    }

    /**
     * 删除临时文件
     *
     * @param paths paths
     */
    public static void deleteTempFiles(List<Path> paths) {
        if (Objects.isNull(paths)) {
            return;
        }
        for (Path path : paths) {
            deleteTempFile(path);
        }
    }

    /**
     * 删除临时文件
     *
     * @param path path
     */
    public static void deleteTempFile(Path path) {
        if (Objects.isNull(path)) {
            return;
        }
        try {
            if (Files.exists(path)) {
                boolean delSuccess = Files.deleteIfExists(path);
                if (!delSuccess) {
                    log.warn("Delete temp file failure,fileName:{}", path.toFile().getName());
                }
            }
        } catch (IOException e) {
            log.warn("Delete temp file failure", e);
        }
    }

    public static Path convertToFile(InputStream is) {
        Path tempFile = createTempFile("i_t", Constants.XLSX);
        try (FileOutputStream fos = new FileOutputStream(tempFile.toFile())) {
            byte[] buffer = new byte[8 * 1024];
            int len;
            while ((len = is.read(buffer)) > 0) {
                fos.write(buffer, 0, len);
            }
        } catch (IOException e) {
            TempFileOperator.deleteTempFile(tempFile);
            throw new SaxReadException("Fail to convert file inputStream to temp file", e);
        }
        return tempFile;
    }

}

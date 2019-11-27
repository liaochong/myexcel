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
package com.github.liaochong.myexcel.core;

import com.github.liaochong.myexcel.utils.TempFileOperator;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;

/**
 * @author liaochong
 * @version 1.0
 */
public class Csv {

    /**
     * csv文件路径
     */
    private Path filePath;

    Csv(Path filePath) {
        this.filePath = filePath;
        byte[] bom = {(byte) 0xEF, (byte) 0xBB, (byte) 0xBF};
        try {
            Files.write(this.filePath, bom);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    public Path getFilePath() {
        return filePath;
    }

    public void write(Path target) {
        this.write(target, false);
    }

    public void write(Path target, boolean append) {
        try {
            if (!append || Files.notExists(target)) {
                Files.createFile(target);
            }
            try (FileInputStream fis = new FileInputStream(filePath.toFile());
                 FileOutputStream fos = new FileOutputStream(target.toFile(), true)) {
                if (append && Files.exists(target) && Files.size(target) > 0) {
                    fis.skip(3);
                }
                byte[] buffer = new byte[8 * 1024];
                int len;
                while ((len = fis.read(buffer)) != -1) {
                    fos.write(buffer, 0, len);
                }
            }
        } catch (IOException e) {
            throw new RuntimeException(e);
        } finally {
            clear();
        }
    }

    public void clear() {
        TempFileOperator.deleteTempFile(filePath);
    }
}

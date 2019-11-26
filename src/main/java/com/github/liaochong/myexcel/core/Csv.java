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

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardOpenOption;

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
    }

    public Path getFilePath() {
        return filePath;
    }

    public void write(Path target) {
        this.write(target, false);
    }

    public void write(Path target, boolean append) {
        Path origin = filePath;
        try {
            if (!Files.exists(target) || Files.size(target) == 0) {
                byte[] result = this.getBomBytes(origin);
                Files.write(target, result);
                return;
            }
            if (append) {
                Files.write(target, Files.readAllBytes(origin), StandardOpenOption.APPEND);
            } else {
                Files.write(target, Files.readAllBytes(origin));
            }
        } catch (IOException e) {
            throw new RuntimeException(e);
        } finally {
            clear();
        }
    }

    private byte[] getBomBytes(Path origin) throws IOException {
        byte[] bom = {(byte) 0xEF, (byte) 0xBB, (byte) 0xBF};
        byte[] originBytes = Files.readAllBytes(origin);
        if (originBytes.length == 0) {
            return bom;
        }
        byte[] result = new byte[bom.length + originBytes.length];
        System.arraycopy(bom, 0, result, 0, bom.length);
        System.arraycopy(originBytes, 0, result, bom.length, originBytes.length);
        return result;
    }

    public void clear() {
        TempFileOperator.deleteTempFile(filePath);
    }
}

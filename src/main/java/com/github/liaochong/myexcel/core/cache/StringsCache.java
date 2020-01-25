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
package com.github.liaochong.myexcel.core.cache;

import com.github.liaochong.myexcel.utils.TempFileOperator;

import java.io.FileInputStream;
import java.io.IOException;
import java.nio.MappedByteBuffer;
import java.nio.channels.FileChannel;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardOpenOption;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

/**
 * Only for xlsx strings cache
 *
 * @author liaochong
 * @version 1.0
 */
public class StringsCache implements Cache<Integer, String> {

    private static final String LINE_SEPARATOR = java.security.AccessController.doPrivileged(
            new sun.security.action.GetPropertyAction("line.separator"));

    private static final int MAX_SIZE_PATH = 1000;

    private static final int MAX_PATH = 5;

    private List<Path> cacheFiles = new ArrayList<>();

    private LinkedHashMap<Integer, List<String>> activeCache = new LinkedHashMap<Integer, List<String>>(MAX_PATH, 0.75F, true) {
        @Override
        protected boolean removeEldestEntry(Map.Entry eldest) {
            return size() > MAX_PATH;
        }
    };

    private List<String> cacheValues = new ArrayList<>(MAX_SIZE_PATH);

    private int totalCount;

    private int activeIndex;

    private int numberOfCacheFile;

    public void init(int totalCount) {
        this.totalCount = totalCount;
        this.numberOfCacheFile = totalCount == MAX_SIZE_PATH ? 1 : totalCount / MAX_SIZE_PATH + 1;
    }

    @Override
    public void cache(Integer key, String value) {
        cacheValues.add(value);
        if ((key + 1) % MAX_SIZE_PATH == 0) {
            if (numberOfCacheFile != 1) {
                writeToFile();
            }
            if (activeIndex == 0) {
                activeCache.put(0, cacheValues);
            }
            boolean isLast = activeIndex == numberOfCacheFile - 1;
            activeIndex++;
            if (isLast) {
                cacheValues = new ArrayList<>(totalCount - numberOfCacheFile * MAX_SIZE_PATH);
            } else {
                cacheValues = new ArrayList<>(MAX_SIZE_PATH);
            }
        }
    }

    @Override
    public String get(Integer key) {
        int route = key / MAX_SIZE_PATH;
        List<String> strings = activeCache.get(route);
        if (strings == null) {
            strings = this.getStrings(route);
            activeCache.put(route, strings);
        }
        return strings.get(key - (route * MAX_SIZE_PATH));
    }

    private List<String> getStrings(int route) {
        Path file = cacheFiles.get(route);
        try (FileInputStream fis = new FileInputStream(file.toFile());
             FileChannel fc = fis.getChannel()) {
            MappedByteBuffer mbb = fc.map(FileChannel.MapMode.READ_ONLY, 0, fc.size());
            byte[] bb = new byte[(int) fc.size()];
            mbb.get(bb);
            String result = new String(bb, StandardCharsets.UTF_8);
            return Arrays.asList(result.split(LINE_SEPARATOR));
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    public void finished() {
        if (cacheValues.isEmpty()) {
            return;
        }
        writeToFile();
    }

    private void writeToFile() {
        Path file = TempFileOperator.createTempFile("s_c", ".data");
        cacheFiles.add(file);
        try {
            Files.write(file, cacheValues, StandardOpenOption.WRITE);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    @Override
    public void clearAll() {
        TempFileOperator.deleteTempFiles(cacheFiles);
    }
}

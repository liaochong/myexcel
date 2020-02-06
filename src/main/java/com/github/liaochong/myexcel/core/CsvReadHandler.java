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

import com.github.liaochong.myexcel.core.constant.Constants;
import com.github.liaochong.myexcel.exception.StopReadException;
import lombok.extern.slf4j.Slf4j;

import java.io.BufferedReader;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.util.List;
import java.util.regex.Pattern;

/**
 * @author liaochong
 * @version 1.0
 */
@Slf4j
class CsvReadHandler<T> extends AbstractReadHandler<T> {

    private static final Pattern PATTERN_SPLIT = Pattern.compile(",(?=([^\\\"]*\\\"[^\\\"]*\\\")*[^\\\"]*$)");

    private static final Pattern PATTERN_QUOTES = Pattern.compile("[\"]{2}");

    private InputStream is;

    private String charset;

    public CsvReadHandler(InputStream is,
                          SaxExcelReader.ReadConfig<T> readConfig,
                          List<T> result) {
        super(true);
        this.is = is;
        this.charset = readConfig.getCharset();
        this.init(result, readConfig);
    }

    public void read() {
        if (is == null) {
            return;
        }
        long startTime = System.currentTimeMillis();
        try (BufferedReader bufferedReader = new BufferedReader(new InputStreamReader(is, charset))) {
            int lineIndex = 0;
            String line;
            while ((line = bufferedReader.readLine()) != null) {
                newRow(lineIndex);
                if (lineIndex == 0 && line.length() >= 1 && line.charAt(0) == '\uFEFF') {
                    line = line.substring(1);
                }
                this.process(line);
                this.initFieldMap();
                lineIndex++;
            }
            log.info("Sax import takes {} ms", System.currentTimeMillis() - startTime);
        } catch (StopReadException e) {
            log.info("Sax import takes {} ms", System.currentTimeMillis() - startTime);
            throw e;
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }

    private void process(String line) {
        if (line != null) {
            String[] strArr = PATTERN_SPLIT.split(line, -1);
            for (int i = 0, size = strArr.length; i < size; i++) {
                String content = strArr[i];
                if (content != null && content.isEmpty()) {
                    content = null;
                }
                if (content != null && content.indexOf(Constants.QUOTES) == 0) {
                    if (content.length() > 2) {
                        content = content.substring(1, content.length() - 1);
                    } else {
                        content = "";
                    }
                }
                if (content != null) {
                    content = PATTERN_QUOTES.matcher(content).replaceAll("\"");
                }
                content = readConfig.getTrim().apply(content);
                this.addTitleConsumer.accept(content, i);
                handleField(i, content);
            }
        }
        handleResult();
    }
}

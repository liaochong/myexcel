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

import com.github.liaochong.myexcel.exception.StopReadException;
import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVParser;
import org.apache.commons.csv.CSVRecord;
import org.apache.commons.io.input.BOMInputStream;
import org.slf4j.Logger;

import java.io.InputStream;
import java.io.InputStreamReader;
import java.io.Reader;
import java.util.Iterator;
import java.util.List;

/**
 * @author liaochong
 * @version 1.0
 */
class CsvReadHandler<T> extends AbstractReadHandler<T> {
    private static final Logger log = org.slf4j.LoggerFactory.getLogger(CsvReadHandler.class);

    private final InputStream is;

    public CsvReadHandler(InputStream is,
                          SaxExcelReader.ReadConfig<T> readConfig,
                          List<T> result) {
        super(true, result, readConfig);
        this.is = is;
    }

    public void read() {
        if (is == null) {
            return;
        }
        long startTime = System.currentTimeMillis();
        try (Reader reader = new InputStreamReader(new BOMInputStream(is), readConfig.csvCharset);
             CSVParser parser = new CSVParser(reader, CSVFormat.EXCEL.withDelimiter(readConfig.csvDelimiter))) {
            for (final CSVRecord record : parser) {
                newRow((int) (record.getRecordNumber() - 1), true);
                Iterator<String> iterator = record.stream().iterator();
                int columnIndex = 0;
                while (iterator.hasNext()) {
                    String content = iterator.next();
                    handleField(columnIndex++, content);
                }
                handleResult();
            }
            log.info("Sax import takes {} ms", System.currentTimeMillis() - startTime);
        } catch (StopReadException e) {
            log.info("Sax import takes {} ms", System.currentTimeMillis() - startTime);
            throw e;
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }
}

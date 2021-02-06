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
import com.github.liaochong.myexcel.core.container.Pair;
import com.github.liaochong.myexcel.core.reflect.ClassFieldContainer;
import com.github.liaochong.myexcel.exception.CsvBuildException;
import com.github.liaochong.myexcel.utils.ReflectUtil;
import com.github.liaochong.myexcel.utils.TempFileOperator;

import java.io.Closeable;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardOpenOption;
import java.util.LinkedList;
import java.util.List;
import java.util.Map;
import java.util.regex.Pattern;
import java.util.stream.Collectors;

/**
 * CSV文件构建器
 *
 * @author liaochong
 * @version 1.0
 */
public class CsvBuilder<T> extends AbstractSimpleExcelBuilder implements Closeable {

    private static final Pattern PATTERN_QUOTES_PREMISE = Pattern.compile("[,\"]+");

    private static final Pattern PATTERN_QUOTES = Pattern.compile("\"");

    /**
     * 文件路径
     */
    private volatile Csv csv;

    private CsvBuilder() {
        super(true);
    }

    public static <T> CsvBuilder<T> of(Class<T> clazz) {
        CsvBuilder<T> csvBuilder = new CsvBuilder<>();
        csvBuilder.isMapBuild = clazz == Map.class;
        if (!csvBuilder.isMapBuild) {
            ClassFieldContainer classFieldContainer = ReflectUtil.getAllFieldsOfClass(clazz);
            csvBuilder.filteredFields = csvBuilder.getFilteredFields(classFieldContainer);
        }
        return csvBuilder;
    }

    public CsvBuilder<T> groups(Class<?>... groups) {
        filteredFields = this.getGroupFields(filteredFields, groups);
        return this;
    }

    public CsvBuilder<T> fieldDisplayOrder(List<String> fieldDisplayOrder) {
        this.fieldDisplayOrder = fieldDisplayOrder;
        return this;
    }

    public CsvBuilder<T> titles(List<String> titles) {
        this.titles = titles;
        return this;
    }

    public CsvBuilder<T> noTitles() {
        this.titles = null;
        return this;
    }

    public Csv build(List<T> beans) {
        return this.doWrite(beans);
    }

    public void append(List<T> beans) {
        this.doWrite(beans);
    }

    public Csv build() {
        return csv;
    }

    private Csv doWrite(List<T> beans) {
        try {
            if (beans == null || beans.isEmpty()) {
                return csv;
            }
            List<List<?>> contents = getRenderContent(beans);
            this.writeToCsv(contents);
        } catch (Exception e) {
            if (csv != null) {
                TempFileOperator.deleteTempFile(csv.getFilePath());
            }
            throw new CsvBuildException("Build csv failure", e);
        }
        return csv;
    }

    /**
     * 获取需要被渲染的内容
     *
     * @param data 数据集合
     * @return 结果集
     */
    @SuppressWarnings("unchecked")
    private List<List<?>> getRenderContent(List<T> data) {
        List<List<?>> result = new LinkedList<>();
        if (isMapBuild) {
            for (T datum : data) {
                List<Pair<? extends Class, ?>> resolvedDataList = this.assemblingMapContents((Map<String, Object>) datum);
                this.appendContent(result, resolvedDataList);
            }
        } else if (hasMultiColumn) {
            for (T datum : data) {
                List<List<Pair<? extends Class, ?>>> contents = this.getMultiRenderContent(datum, filteredFields);
                for (List<Pair<? extends Class, ?>> content : contents) {
                    this.appendContent(result, content);
                }
            }
        } else {
            for (T datum : data) {
                List<Pair<? extends Class, ?>> contents = this.getOriginalRenderContent(datum, filteredFields);
                this.appendContent(result, contents);
            }
        }
        return result;
    }

    private void appendContent(List<List<?>> result, List<Pair<? extends Class, ?>> resolvedDataList) {
        List<?> values = resolvedDataList.stream().map(Pair::getValue).collect(Collectors.toList());
        result.add(values);
    }

    private void writeToCsv(List<List<?>> data) {
        if (titles != null) {
            synchronized (this) {
                if (titles == null) {
                    return;
                }
                data.add(0, titles);
                titles = null;
            }
        }
        List<String> content = data.stream().map(d -> {
            return d.stream().map(v -> {
                if (v == null) {
                    return "";
                }
                String vStr = v.toString();
                vStr = PATTERN_QUOTES.matcher(vStr).replaceAll("\"\"");
                boolean hasComma = PATTERN_QUOTES_PREMISE.matcher(v.toString()).find();
                if (hasComma) {
                    vStr = "\"" + vStr + "\"";
                }
                return vStr;
            }).collect(Collectors.joining(Constants.COMMA));
        }).collect(Collectors.toCollection(LinkedList::new));

        synchronized (this) {
            try {
                if (csv == null) {
                    Path csvTemp = TempFileOperator.createTempFile("d_t_c", Constants.CSV);
                    csv = new Csv(csvTemp);
                }
                Files.write(csv.getFilePath(), content, StandardCharsets.UTF_8, StandardOpenOption.APPEND);
            } catch (IOException e) {
                throw new RuntimeException(e);
            }
        }
    }

    @Override
    public void close() throws IOException {
        clear();
    }

    public void clear() {
        if (csv != null) {
            csv.clear();
        }
    }
}

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

import com.github.liaochong.myexcel.core.annotation.ExcelColumn;
import com.github.liaochong.myexcel.core.constant.Constants;
import com.github.liaochong.myexcel.core.container.Pair;
import com.github.liaochong.myexcel.core.container.ParallelContainer;
import com.github.liaochong.myexcel.core.reflect.ClassFieldContainer;
import com.github.liaochong.myexcel.exception.CsvBuildException;
import com.github.liaochong.myexcel.utils.ReflectUtil;
import com.github.liaochong.myexcel.utils.StringUtil;
import com.github.liaochong.myexcel.utils.TempFileOperator;

import java.io.Closeable;
import java.io.IOException;
import java.lang.reflect.Field;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardOpenOption;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.HashMap;
import java.util.LinkedList;
import java.util.List;
import java.util.regex.Pattern;
import java.util.stream.Collectors;
import java.util.stream.IntStream;

/**
 * CSV文件构建器
 *
 * @author liaochong
 * @version 1.0
 */
public class CsvBuilder<T> extends AbstractSimpleExcelBuilder implements Closeable {

    private static final Pattern PATTERN_QUOTES_PREMISE = Pattern.compile("[,\"]+");

    private static final Pattern PATTERN_QUOTES = Pattern.compile("\"");

    private List<Field> fields;
    /**
     * 文件路径
     */
    private volatile Csv csv;

    private CsvBuilder() {
    }

    public static <T> CsvBuilder<T> of(Class<T> clazz) {
        CsvBuilder<T> csvBuilder = new CsvBuilder<>();
        ClassFieldContainer classFieldContainer = ReflectUtil.getAllFieldsOfClass(clazz);
        csvBuilder.fields = csvBuilder.getFields(classFieldContainer);
        return csvBuilder;
    }

    public CsvBuilder<T> groups(Class<?>... groups) {
        fields = this.getGroupFields(fields, groups);
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

    private List<Field> getFields(ClassFieldContainer classFieldContainer, Class<?>... groups) {
        GlobalSetting globalSetting = new GlobalSetting();
        setGlobalSetting(classFieldContainer, globalSetting);

        List<Field> preElectionFields = this.getPreElectionFields(classFieldContainer, globalSetting);
        List<Field> sortedFields = getGroupFields(preElectionFields, groups);
        List<String> titles = new ArrayList<>(preElectionFields.size());
        defaultValueMap = new HashMap<>(preElectionFields.size());
        boolean needToAddTitle = this.titles == null;
        for (Field field : sortedFields) {
            ExcelColumn excelColumn = field.getAnnotation(ExcelColumn.class);
            if (excelColumn != null) {
                if (needToAddTitle) {
                    if (globalSetting.isUseFieldNameAsTitle() && excelColumn.title().isEmpty()) {
                        titles.add(field.getName());
                    } else {
                        titles.add(excelColumn.title());
                    }
                }
                if (!excelColumn.defaultValue().isEmpty()) {
                    defaultValueMap.put(field, excelColumn.defaultValue());
                }
            } else {
                if (needToAddTitle) {
                    if (globalSetting.isUseFieldNameAsTitle()) {
                        titles.add(field.getName());
                    } else {
                        titles.add(null);
                    }
                }
            }
        }
        boolean hasTitle = titles.stream().anyMatch(StringUtil::isNotBlank);
        if (hasTitle) {
            this.titles = titles;
        }
        return sortedFields;
    }

    /**
     * 获取需要被渲染的内容
     *
     * @param data 数据集合
     * @return 结果集
     */
    private List<List<?>> getRenderContent(List<T> data) {
        List<ParallelContainer> resolvedDataContainers = IntStream.range(0, data.size()).parallel().mapToObj(index -> {
            List<?> resolvedDataList = this.getRenderContent(data.get(index), fields);
            return new ParallelContainer<>(index, resolvedDataList);
        }).collect(Collectors.toCollection(LinkedList::new));
        data.clear();

        // 重排序
        return resolvedDataContainers.stream()
                .sorted(Comparator.comparing(ParallelContainer::getIndex))
                .map(ParallelContainer<List<Pair<? extends Class, ?>>>::getData).collect(Collectors.toCollection(LinkedList::new));
    }

    private void writeToCsv(List<List<?>> data) {
        if (titles != null) {
            data.add(0, titles);
            titles = null;
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

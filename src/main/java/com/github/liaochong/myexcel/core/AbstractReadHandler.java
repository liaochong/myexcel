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
import com.github.liaochong.myexcel.core.converter.ConvertContext;
import com.github.liaochong.myexcel.core.converter.ReadConverterContext;
import com.github.liaochong.myexcel.core.reflect.ClassFieldContainer;
import com.github.liaochong.myexcel.exception.StopReadException;
import com.github.liaochong.myexcel.utils.ConfigurationUtil;
import com.github.liaochong.myexcel.utils.ReflectUtil;
import com.github.liaochong.myexcel.utils.StringUtil;

import java.lang.reflect.Field;
import java.util.Collections;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.StringJoiner;
import java.util.function.BiConsumer;
import java.util.function.Consumer;
import java.util.function.Supplier;
import java.util.stream.Collectors;

/**
 * 读取抽象
 *
 * @author liaochong
 * @version 1.0
 */
abstract class AbstractReadHandler<T> {

    private Map<Integer, Field> fieldMap;

    private T obj;

    protected Map<Integer, Map<Integer, String>> titles = new LinkedHashMap<>();

    protected SaxExcelReader.ReadConfig<T> readConfig;

    private final ReadContext<T> context = new ReadContext<>();

    private final RowContext rowContext = new RowContext();

    private final ConvertContext convertContext;
    /**
     * Row object currently being processed
     */
    private final Row currentRow = new Row(-1);
    /**
     * 上一列列号
     */
    private int prevColNum = -1;

    private Supplier<T> newInstance;

    private BiConsumer<Integer, String> fieldHandler;

    private Consumer<T> resultHandler;

    private BiConsumer<T, RowContext> contextResultHandler;

    /**
     * Whether to use title for import
     */
    private boolean readWithTitle;
    /**
     * 标题行编号，默认为-1
     */
    private int titleRowNum = -1;

    /**
     * is blank row
     */
    protected boolean isBlankRow;

    public AbstractReadHandler(boolean readCsv,
                               List<T> result,
                               SaxExcelReader.ReadConfig<T> readConfig) {
        convertContext = new ConvertContext(readCsv);
        Class<T> dataType = readConfig.dataType;
        fieldMap = ReflectUtil.getFieldMapOfExcelColumn(dataType);
        this.readConfig = readConfig;
        boolean isMapType = dataType == Map.class;
        readWithTitle = !isMapType && fieldMap.isEmpty();
        setNewInstanceFunction(dataType, isMapType);
        // 全局配置获取
        setConfiguration(dataType, isMapType);
        setResultHandlerFunction(result, readConfig);
        setFieldHandlerFunction(isMapType);
    }

    private void setResultHandlerFunction(List<T> result, SaxExcelReader.ReadConfig<T> readConfig) {
        if (readConfig.consumer != null) {
            resultHandler = v -> readConfig.consumer.accept(v);
        } else if (readConfig.function != null) {
            resultHandler = v -> {
                Boolean noStop = readConfig.function.apply(v);
                if (!noStop) {
                    throw new StopReadException();
                }
            };
        } else if (readConfig.contextConsumer != null) {
            contextResultHandler = (v, context) -> readConfig.contextConsumer.accept(v, context);
        } else if (readConfig.contextFunction != null) {
            contextResultHandler = (v, context) -> {
                Boolean noStop = readConfig.contextFunction.apply(v, context);
                if (!noStop) {
                    throw new StopReadException();
                }
            };
        } else {
            resultHandler = result::add;
        }
    }

    @SuppressWarnings("unchecked")
    private void setNewInstanceFunction(Class<T> dataType, boolean isMapType) {
        if (isMapType) {
            newInstance = () -> (T) new LinkedHashMap<Cell, String>();
        } else {
            newInstance = () -> ReflectUtil.newInstance(dataType);
        }
    }

    private void setConfiguration(Class<T> dataType, boolean isMapType) {
        if (isMapType) {
            return;
        }
        ClassFieldContainer classFieldContainer = ReflectUtil.getAllFieldsOfClass(dataType);
        ConfigurationUtil.parseConfiguration(classFieldContainer, convertContext.configuration);

        List<Field> fields = classFieldContainer.getFieldsByAnnotation(ExcelColumn.class);
        fields.forEach(field -> {
            ExcelColumn excelColumn = field.getAnnotation(ExcelColumn.class);
            if (excelColumn == null) {
                return;
            }
            ExcelColumnMapping mapping = ExcelColumnMapping.mapping(excelColumn);
            convertContext.excelColumnMappingMap.put(field, mapping);
        });
    }

    @SuppressWarnings("unchecked")
    private void setFieldHandlerFunction(boolean isMapType) {
        if (isMapType) {
            fieldHandler = (colNum, content) -> {
                for (int i = prevColNum + 1; i < colNum; i++) {
                    ((Map<Cell, String>) obj).put(new Cell(currentRow.getRowNum(), i), null);
                }
                ((Map<Cell, String>) obj).put(new Cell(currentRow.getRowNum(), colNum), content);
                prevColNum = colNum;
            };
        } else {
            fieldHandler = (colNum, content) -> {
                Field field = fieldMap.get(colNum);
                convert(content, currentRow.getRowNum(), colNum, field);
            };
        }
    }

    protected void convert(String value, int rowNum, int colNum, Field field) {
        if (value == null || field == null) {
            return;
        }
        context.reset(obj, field, value, rowNum, colNum);
        ReadConverterContext.convert(obj, context, convertContext, readConfig.exceptionFunction);
    }

    protected void newRow(int rowNum) {
        currentRow.setRowNum(rowNum);
        obj = newInstance.get();
        prevColNum = -1;
        isBlankRow = true;
    }

    protected void setRecordAsNull() {
        obj = null;
    }

    protected void handleField(Integer colNum, String content) {
        if (obj == null || colNum < 0) {
            return;
        }
        isBlankRow = false;
        content = readConfig.trim.apply(content);
        if (readConfig.rowFilter.test(currentRow)) {
            fieldHandler.accept(colNum, content);
        } else if (readWithTitle) {
            Map<Integer, String> rowMapping = titles.computeIfAbsent(currentRow.getRowNum(), rowNum -> new HashMap<>());
            rowMapping.put(colNum, content);
            if (titleRowNum == -1) {
                // 尝试下一行是否为标题行
                Row nextRow = new Row(currentRow.getRowNum() + 1);
                if (readConfig.rowFilter.test(nextRow)) {
                    titleRowNum = currentRow.getRowNum();
                }
            }
        }
    }

    protected void handleResult() {
        if (isBlankRow) {
            if (readConfig.stopReadingOnBlankRow) {
                throw new StopReadException();
            }
            // 忽略空白行
            if (readConfig.ignoreBlankRow) {
                return;
            }
        }
        this.initFieldMap();
        if (!readConfig.rowFilter.test(currentRow)) {
            return;
        }
        if (!readConfig.beanFilter.test(obj)) {
            return;
        }
        if (readWithTitle && currentRow.getRowNum() == 0) {
            readWithTitle = false;
            return;
        }
        if (resultHandler != null) {
            resultHandler.accept(obj);
        } else {
            rowContext.setRowNum(currentRow.getRowNum());
            contextResultHandler.accept(obj, rowContext);
        }
    }

    private void initFieldMap() {
        if (currentRow.getRowNum() != titleRowNum || !fieldMap.isEmpty()) {
            return;
        }
        Map<String, Field> titleFieldMap = ReflectUtil.getFieldMapOfTitleExcelColumn(readConfig.dataType);
        fieldMap = new HashMap<>(titleFieldMap.size());
        // 获取最大列数
        List<Integer> colNums = titles.values().stream().flatMap(t -> t.keySet().stream()).collect(Collectors.toList());
        int maxColNum = Collections.max(colNums);
        // 获取最终标题行
        Map<Integer, String> titleMapping = titles.get(titleRowNum);
        for (int i = 0; i <= maxColNum; i++) {
            StringJoiner realTitle = new StringJoiner(Constants.ARROW);
            int colNum = i;
            titles.keySet().forEach(rowNum -> {
                if (rowNum == titleRowNum) {
                    return;
                }
                Map<Integer, String> prevColMapping = titles.get(rowNum);
                int realColNum = colNum;
                for (; ; ) {
                    String prevTitle = prevColMapping.get(realColNum);
                    if (StringUtil.isNotBlank(prevTitle)) {
                        realTitle.add(prevTitle);
                        return;
                    }
                    realColNum -= 1;
                }
            });
            final String title = titleMapping.get(i);
            if (StringUtil.isNotBlank(title)) {
                realTitle.add(title);
            }
            fieldMap.put(colNum, titleFieldMap.get(realTitle.toString()));
        }
        // 释放
        titles = null;
    }
}

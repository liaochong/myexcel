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
package com.github.liaochong.myexcel.core;

/**
 * @author liaochong
 * @version 1.0
 */
public enum WorkbookType {
    /**
     * none,default xlsx
     */
    NONE,
    /**
     * .xls
     */
    XLS,
    /**
     * .xlsx
     */
    XLSX,
    /**
     * .xlsx
     */
    SXLSX;

    public static boolean isXls(WorkbookType workbookType) {
        return XLS.equals(workbookType);
    }

    public static boolean isSxlsx(WorkbookType workbookType) {
        return SXLSX.equals(workbookType);
    }

    public static boolean isNone(WorkbookType workbookType) {
        return NONE.equals(workbookType);
    }
}

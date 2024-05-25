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
package com.github.liaochong.myexcel.utils;

import org.hibernate.validator.HibernateValidator;

import javax.validation.Validation;
import javax.validation.Validator;
import javax.validation.ValidatorFactory;

/**
 * @author liaochong
 * @version 1.0
 */
public final class ValidatorUtil {

    private static Validator validator;

    public static synchronized Validator getValidator() {
        if (validator == null) {
            try (ValidatorFactory validatorFactory = Validation
                    .byProvider(HibernateValidator.class)
                    .configure()
                    .buildValidatorFactory()) {
                validator = validatorFactory.getValidator();
            }
        }
        return validator;
    }
}

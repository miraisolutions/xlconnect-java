/*
 *
    XLConnect
    Copyright (C) 2010-2025 Mirai Solutions GmbH

    This program is free software: you can redistribute it and/or modify
    it under the terms of the GNU General Public License as published by
    the Free Software Foundation, either version 3 of the License, or
    (at your option) any later version.

    This program is distributed in the hope that it will be useful,
    but WITHOUT ANY WARRANTY; without even the implied warranty of
    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
    GNU General Public License for more details.

    You should have received a copy of the GNU General Public License
    along with this program.  If not, see <http://www.gnu.org/licenses/>.
 *
 */

package com.miraisolutions.xlconnect;

import java.util.*;

/**
 * Wraps a value with attributes to be set on the value in R. @see XLConnect::xlcCall and 
 * https://www.r-bloggers.com/2020/10/attributes-in-r/ for more information.
 */
public class ResultWithAttributes<T> {

    private final T value;

    private final Map<String, String[]> attributes;

    public ResultWithAttributes(T value, Map<String, String[]> theAttributes) {
        this.value = value;
        this.attributes = theAttributes;
    }

    public ResultWithAttributes(T value) {
        this(value, Collections.emptyMap());
    }

    public ResultWithAttributes(T value, Attribute attributeName, String attributeValue) {
        this(value, Collections.singletonMap(attributeName.toString(), new String[] { attributeValue }));
    }

    public T getValue() {
        return value;
    }

    public Map<String, String[]> getAttributes() {
        return Collections.unmodifiableMap(attributes);
    }

    public String[] getAttributeNames() {
        return attributes.keySet().toArray(new String[0]);
    }

    public String[] getAttributeValue(String attributeName) {
        return attributes.get(attributeName);
    }
}

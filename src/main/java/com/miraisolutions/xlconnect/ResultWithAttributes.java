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

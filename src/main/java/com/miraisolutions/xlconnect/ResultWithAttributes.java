package com.miraisolutions.xlconnect;

import java.util.*;

/**
 * Represent attributes to be set on an object in R as part of the result to be returned.
 * Should be extended for each required type. Not using a generic typed value, because it looks like R can't retrieve it
 * in a specific subtype (we get an Object instance).
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

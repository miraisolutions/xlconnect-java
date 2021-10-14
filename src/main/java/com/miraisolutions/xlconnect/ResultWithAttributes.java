package com.miraisolutions.xlconnect;


import java.util.*;

/**
 * Represent attributes to be set on an object in R as part of the result to be returned.
 * Should be extended for each required type. Not using a generic typed value, because it looks like R can't retrieve it
 * in a specific subtype (we get an Object instance).
 */
class ResultWithAttributes {

    private final Map<String,String> attributes;

    public ResultWithAttributes(Map<String,String> theAttributes) {
        this.attributes = theAttributes;
    }

    public ResultWithAttributes(Attribute attributeName, String attributeValue) {
        this(Collections.singletonMap(attributeName.toString(), attributeValue));
    }

    public String[] getAttributeNames() {
        return attributes.keySet().toArray(new String[0]);
    }

    public String[] getAttributeValues() {
        return attributes.values().toArray(new String[0]);
    }

    public String getAttributeValue(String attributeName){
        return attributes.get(attributeName);
    }
}

package com.miraisolutions.xlconnect;

public class BooleanWithAttributes extends ResultWithAttributes {

    private boolean value;

    public BooleanWithAttributes(Attribute attributeName, String attributeValue, boolean value) {
        super(attributeName, attributeValue);
        this.value = value;
    }

    public boolean getValue() {
        return value;
    }
}

package com.miraisolutions.xlconnect;

import com.miraisolutions.xlconnect.data.DataFrame;

public class DataFrameWithAttributes extends ResultWithAttributes{

    private final DataFrame value;

    public DataFrameWithAttributes(Attribute attributeName, String attributeValue, DataFrame value) {
        super(attributeName, attributeValue);
        this.value = value;
    }


    public DataFrame getValue() {
        return value;
    }

}

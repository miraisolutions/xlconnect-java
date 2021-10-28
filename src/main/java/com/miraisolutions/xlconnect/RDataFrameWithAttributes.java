package com.miraisolutions.xlconnect;

import com.miraisolutions.xlconnect.integration.r.RDataFrameWrapper;

import java.util.Map;

public class RDataFrameWithAttributes extends ResultWithAttributes{

    private final RDataFrameWrapper value;

    public RDataFrameWithAttributes(Map<String, String> attributes, RDataFrameWrapper value) {
        super(attributes);
        this.value = value;
    }

    public RDataFrameWithAttributes(RDataFrameWrapper value) {
        super();
        this.value = value;
    }

    public RDataFrameWrapper getValue() {
        return value;
    }
    
}

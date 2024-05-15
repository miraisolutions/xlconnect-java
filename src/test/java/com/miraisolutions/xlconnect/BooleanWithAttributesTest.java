package com.miraisolutions.xlconnect;

import org.junit.Test;

import static org.junit.Assert.*;

public class BooleanWithAttributesTest {

    @Test
    public void getTrueValue() {
        ResultWithAttributes<Boolean> underTest = new ResultWithAttributes<>(true,Attribute.WORKSHEET_SCOPE, "bla");
        assert underTest.getValue();
        assertEquals(1, underTest.getAttributeNames().length);
    }
}
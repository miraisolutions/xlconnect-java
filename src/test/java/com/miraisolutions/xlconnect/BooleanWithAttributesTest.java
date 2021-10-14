package com.miraisolutions.xlconnect;

import org.junit.Before;
import org.junit.Test;

import static org.junit.Assert.*;

public class BooleanWithAttributesTest {

    @Test
    public void getTrueValue() {
        BooleanWithAttributes underTest = new BooleanWithAttributes(Attribute.WORKSHEET_SCOPE, "bla", true);
        assert underTest.getValue();
        assertEquals(1, underTest.getAttributeNames().length);
    }
}
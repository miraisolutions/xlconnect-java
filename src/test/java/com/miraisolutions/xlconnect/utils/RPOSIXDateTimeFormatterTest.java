package com.miraisolutions.xlconnect.utils;

import org.junit.After;
import org.junit.Test;
import org.junit.Before;

import java.time.LocalDateTime;
import java.time.ZoneId;
import java.util.Calendar;
import java.util.Date;
import java.util.TimeZone;

import static org.junit.Assert.*;

public class RPOSIXDateTimeFormatterTest {

    private RPOSIXDateTimeFormatter underTest;

    @Before
    public void beforeEach() {
        underTest = new RPOSIXDateTimeFormatter();
    }

    @Test
    public void parseDate() {
        Date result = underTest.parse("06.02.2012 16:15:23", "%d.%m.%Y %H:%M:%S");
        //s of JDK version 1.1, replaced by Calendar.get(Calendar.YEAR) - 1900.
        Calendar cal = Calendar.getInstance(TimeZone.getDefault());
        cal.setTime(result);
        assertEquals(16, cal.get(Calendar.HOUR_OF_DAY));
        assertEquals(23, cal.get(Calendar.SECOND));
        assertEquals(2012, cal.get(Calendar.YEAR));
    }

    @Test
    public void parseDateInDST() {
        Date result = underTest.parse("06.07.2012 16:15:23", "%d.%m.%Y %H:%M:%S");
        //s of JDK version 1.1, replaced by Calendar.get(Calendar.YEAR) - 1900.
        Calendar cal = Calendar.getInstance(TimeZone.getDefault());
        cal.setTime(result);
        assertEquals(16, cal.get(Calendar.HOUR_OF_DAY));
        assertEquals(23, cal.get(Calendar.SECOND));
        assertEquals(2012, cal.get(Calendar.YEAR));
    }
}
package com.miraisolutions.xlconnect.utils;

import org.junit.Before;
import org.junit.Test;

import java.time.Instant;
import java.time.ZoneId;
import java.time.ZonedDateTime;
import java.util.Calendar;
import java.util.Date;
import java.util.TimeZone;

import static org.junit.Assert.assertEquals;

public class RPOSIXDateTimeFormatterTest {

    private RPOSIXDateTimeFormatter underTest;

    @Before
    public void beforeEach() {
        underTest = new RPOSIXDateTimeFormatter();
    }

    @Test
    public void parseDate() {
        Date result = underTest.parse("06.11.2012 07:15:23", "%d.%m.%Y %H:%M:%S");
        Calendar cal = Calendar.getInstance(TimeZone.getDefault());
        cal.setTime(result);
        assertEquals(7, cal.get(Calendar.HOUR_OF_DAY));
        assertEquals(15, cal.get(Calendar.MINUTE));
        assertEquals(23, cal.get(Calendar.SECOND));
        assertEquals(6, cal.get(Calendar.DAY_OF_MONTH));
        assertEquals(Calendar.NOVEMBER, cal.get(Calendar.MONTH));
        assertEquals(2012, cal.get(Calendar.YEAR));
    }

    @Test
    public void parseDateInDST() {
        Date result = underTest.parse("06.07.2012 16:15:23", "%d.%m.%Y %H:%M:%S");
        Calendar cal = Calendar.getInstance(TimeZone.getDefault());
        cal.setTime(result);
        assertEquals(16, cal.get(Calendar.HOUR_OF_DAY));
        assertEquals(15, cal.get(Calendar.MINUTE));
        assertEquals(23, cal.get(Calendar.SECOND));
        assertEquals(6, cal.get(Calendar.DAY_OF_MONTH));
        assertEquals(Calendar.JULY, cal.get(Calendar.MONTH));
        assertEquals(2012, cal.get(Calendar.YEAR));
    }

    @Test
    public void formatDate() {
        Date input = Date.from(Instant.from(ZonedDateTime.of(2012,2,6,16,15,23,0, ZoneId.systemDefault())));
        String result = underTest.format(input, "%d.%m.%Y %H:%M:%S");
        assertEquals("06.02.2012 16:15:23", result);
    }

    @Test
    public void formatDateInDST() {
        Date input = Date.from(Instant.from(ZonedDateTime.of(2012,7,6,16,15,23,0, ZoneId.systemDefault())));
        String result = underTest.format(input, "%d.%m.%Y %H:%M:%S");
        assertEquals("06.07.2012 16:15:23", result);
    }
}
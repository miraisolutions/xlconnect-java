/*
 *
    XLConnect
    Copyright (C) 2010-2025 Mirai Solutions GmbH

    This program is free software: you can redistribute it and/or modify
    it under the terms of the GNU General Public License as published by
    the Free Software Foundation, either version 3 of the License, or
    (at your option) any later version.

    This program is distributed in the hope that it will be useful,
    but WITHOUT ANY WARRANTY; without even the implied warranty of
    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
    GNU General Public License for more details.

    You should have received a copy of the GNU General Public License
    along with this program.  If not, see <http://www.gnu.org/licenses/>.
 *
 */

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

public final class RPOSIXDateTimeFormatterTest {

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
        Date input = Date.from(Instant.from(ZonedDateTime.of(2012, 2, 6, 16, 15, 23, 0, ZoneId.systemDefault())));
        String result = underTest.format(input, "%d.%m.%Y %H:%M:%S");
        assertEquals("06.02.2012 16:15:23", result);
    }

    @Test
    public void formatDateInDST() {
        Date input = Date.from(Instant.from(ZonedDateTime.of(2012, 7, 6, 16, 15, 23, 0, ZoneId.systemDefault())));
        String result = underTest.format(input, "%d.%m.%Y %H:%M:%S");
        assertEquals("06.07.2012 16:15:23", result);
    }
}
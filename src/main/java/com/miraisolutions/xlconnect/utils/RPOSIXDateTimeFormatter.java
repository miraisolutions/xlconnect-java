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

import java.time.Instant;
import java.time.LocalDateTime;
import java.time.ZoneId;
import java.time.ZonedDateTime;
import java.time.format.DateTimeFormatterBuilder;
import java.time.format.DateTimeParseException;
import java.time.format.TextStyle;
import java.time.temporal.ChronoField;
import java.time.temporal.TemporalAccessor;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;

/**
 * Inspired by <a href="http://code.google.com/p/renjin/source/browse/trunk/core/src/main/java/r/base/Time.java?spec=svn379&r=379">Time.java</a>
 */
public final class RPOSIXDateTimeFormatter implements DateTimeFormatter {

    private final Map<String, java.time.format.DateTimeFormatter> cache = new HashMap<>();

    private java.time.format.DateTimeFormatter getFormatter(String format) {
        return cache.computeIfAbsent(format, RPOSIXDateTimeFormatter::buildFormatter);
    }

    private static java.time.format.DateTimeFormatter buildFormatter(String format) {
        DateTimeFormatterBuilder builder = new DateTimeFormatterBuilder();

        for (int i = 0; i < format.length(); ++i) {
            if (format.charAt(i) == '%' && i + 1 < format.length()) {
                char specifier = format.charAt(++i);
                switch (specifier) {
                    case '%' -> builder.appendLiteral("%");
                    // Abbreviated weekday name in the current locale. (Also matches
                    // full name on input.)
                    case 'a' -> builder.appendText(ChronoField.DAY_OF_WEEK, TextStyle.SHORT);
                    // Full weekday name in the current locale.  (Also matches
                    // abbreviated name on input.)
                    case 'A' -> builder.appendText(ChronoField.DAY_OF_WEEK, TextStyle.FULL);
                    // Abbreviated month name in the current locale. (Also matches
                    // full name on input.)
                    case 'b' -> builder.appendText(ChronoField.MONTH_OF_YEAR, TextStyle.SHORT);
                    // Full month name in the current locale.  (Also matches
                    // abbreviated name on input.)
                    case 'B' -> builder.appendText(ChronoField.MONTH_OF_YEAR, TextStyle.FULL);
                    //  Date and time.  Locale-specific on output, �"%a %b %e
                    // %H:%M:%S %Y"� on input.
                    case 'c' -> throw new UnsupportedOperationException("%c not yet implemented");
                    // Day of the month as decimal number (01-31).
                    case 'd' -> builder.appendValue(ChronoField.DAY_OF_MONTH, 2);
                    // Hours as decimal number (00-23).
                    case 'H' -> builder.appendValue(ChronoField.HOUR_OF_DAY, 2);
                    // Hours as decimal number (01-12).
                    case 'I' -> builder.appendValue(ChronoField.CLOCK_HOUR_OF_AMPM, 2);
                    // Day of year as decimal number (001-366).
                    case 'j' -> builder.appendValue(ChronoField.DAY_OF_YEAR, 3);
                    // Month as decimal number (01-12).
                    case 'm' -> builder.appendValue(ChronoField.MONTH_OF_YEAR, 2);
                    // Minute as decimal number (00-59).
                    case 'M' -> builder.appendValue(ChronoField.MINUTE_OF_HOUR, 2);
                    // AM/PM indicator in the locale.  Used in conjunction with �%I�
                    // and *not* with �%H�.  An empty string in some locales.
                    case 'p' -> builder.appendText(ChronoField.AMPM_OF_DAY);
                    case 'O' -> {
                        if (i + 1 >= format.length()) {
                            builder.appendLiteral("%O");
                        } else {
                            if (format.charAt(++i) == 'S') {
                                int n = 3;
                                if (i + 1 < format.length()) {
                                    n = Integer.parseInt(Character.toString(format.charAt(++i)));
                                }

                                builder.appendValue(ChronoField.SECOND_OF_MINUTE);
                                if (n > 0) {
                                    builder.appendLiteral('.');
                                    builder.appendFraction(ChronoField.SECOND_OF_MINUTE, n, n, true);
                                }
                            } else {
                                throw new UnsupportedOperationException("%O[dHImMUVwWy] not yet implemented");
                            }
                        }
                    }
                    // Second as decimal number (00-61), allowing for up to two
                    // leap-seconds (but POSIX-compliant implementations will ignore
                    // leap seconds).
                    // TODO: I have no idea what the docs are talking about in relation
                    // to leap seconds
                    case 'S' -> builder.appendValue(ChronoField.SECOND_OF_MINUTE, 2);
                    // case 'U':
                    // Week of the year as decimal number (00-53) using Sunday as
                    // the first day 1 of the week (and typically with the first
                    //  Sunday of the year as day 1 of week 1).  The US convention.
                    // case 'w':
                    // Weekday as decimal number (0-6, Sunday is 0).

                    // case 'W':
                    // Week of the year as decimal number (00-53) using Monday as
                    // the first day of week (and typically with the first Monday of
                    // the year as day 1 of week 1). The UK convention.

                    // �%x� Date.  Locale-specific on output, �"%y/%m/%d"� on input.


                    //�%X� Time.  Locale-specific on output, �"%H:%M:%S"� on input.

                    // Year without century (00-99). Values 00 to 68 are prefixed by
                    // 20 and 69 to 99 by 19 - that is the behaviour specified by
                    // the 2004 POSIX standard, but it does also say �it is expected
                    // that in a future version the default century inferred from a
                    // 2-digit year will change�.
                    case 'y' -> builder.appendValueReduced(ChronoField.YEAR, 2, 2, 1969);
                    // Year with century
                    case 'Y' -> builder.appendValue(ChronoField.YEAR, 4);
                    // Signed offset in hours and minutes from UTC, so �-0800� is 8
                    // hours behind UTC.
                    case 'z' -> builder.appendOffset("+HH:mm", "+0000");
                    // (output only.) Time zone as a character string (empty if not
                    // available).
                    case 'Z' -> builder.appendZoneOrOffsetId();
                    default -> throw new UnsupportedOperationException("%" + specifier + " not yet implemented");
                }
            } else {
                builder.appendLiteral(format.substring(i, i + 1));
            }
        }
        return builder.toFormatter();
    }

    public String format(Date d, String format) {
        return getFormatter(format).format(LocalDateTime.ofInstant(d.toInstant(), ZoneId.systemDefault()));
    }

    public Date parse(String s, String format) {
        java.time.format.DateTimeFormatter formatter = getFormatter(format);
        ZonedDateTime zoned;
        try {
            zoned = ZonedDateTime.parse(s, formatter);
        } catch (DateTimeParseException e) {
            TemporalAccessor parsed = formatter.parse(s);
            zoned = ZonedDateTime.of(LocalDateTime.from(parsed), ZoneId.systemDefault());
        }
        return new Date(Instant.from(zoned).toEpochMilli());
    }

}

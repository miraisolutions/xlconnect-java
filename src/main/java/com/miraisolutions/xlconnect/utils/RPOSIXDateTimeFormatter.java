/*
 *
    XLConnect
    Copyright (C) 2010-2018 Mirai Solutions GmbH

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

import java.time.LocalDateTime;
import java.time.OffsetDateTime;
import java.time.ZoneId;
import java.time.ZoneOffset;
import java.time.format.DateTimeFormatterBuilder;
import java.time.format.TextStyle;
import java.time.temporal.ChronoField;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;

/**
 * Inspired by http://code.google.com/p/renjin/source/browse/trunk/core/src/main/java/r/base/Time.java?spec=svn379&r=379
 * @author Martin Studer, Mirai Solutions GmbH
 */
public class RPOSIXDateTimeFormatter implements DateTimeFormatter {

    private Map<String, java.time.format.DateTimeFormatter> cache = new HashMap<>();

    private java.time.format.DateTimeFormatter getFormatter(String format) {
        if(cache.containsKey(format))
            return cache.get(format);
        
        DateTimeFormatterBuilder builder = new DateTimeFormatterBuilder();

        for(int i=0;i<format.length();++i) {
          if(format.charAt(i)=='%' && i+1 < format.length()) {
            char specifier = format.charAt(++i);
            switch(specifier) {
            case '%':
              builder.appendLiteral("%");
              break;
            case 'a':
              // Abbreviated weekday name in the current locale. (Also matches
              // full name on input.)
              builder.appendText(ChronoField.DAY_OF_WEEK, TextStyle.SHORT);
              break;
            case 'A':
              // Full weekday name in the current locale.  (Also matches
              // abbreviated name on input.)
              builder.appendText(ChronoField.DAY_OF_WEEK, TextStyle.FULL);
              break;
            case 'b':
              // Abbreviated month name in the current locale. (Also matches
              // full name on input.)
                builder.appendText(ChronoField.MONTH_OF_YEAR, TextStyle.SHORT);
              break;
            case 'B':
              // Full month name in the current locale.  (Also matches
              // abbreviated name on input.)
                builder.appendText(ChronoField.MONTH_OF_YEAR, TextStyle.FULL);
              break;
            case 'c':
              //  Date and time.  Locale-specific on output, �"%a %b %e
              // %H:%M:%S %Y"� on input.
              throw new UnsupportedOperationException("%c not yet implemented");
            case 'd':
              // Day of the month as decimal number (01-31).
                builder.appendValue(ChronoField.DAY_OF_MONTH,2);
              break;
            case 'H':
              // Hours as decimal number (00-23).
              builder.appendValue(ChronoField.HOUR_OF_DAY, 2);
              break;
            case 'I':
              // Hours as decimal number (01-12).
              builder.appendValue(ChronoField.CLOCK_HOUR_OF_AMPM, 2);
              break;
            case 'j':
              // Day of year as decimal number (001-366).
              builder.appendValue(ChronoField.DAY_OF_YEAR,3);
              break;
            case 'm':
              // Month as decimal number (01-12).
                builder.appendValue(ChronoField.MONTH_OF_YEAR, 2);
              break;
            case 'M':
              // Minute as decimal number (00-59).
                builder.appendValue(ChronoField.MINUTE_OF_HOUR, 2);
              break;
            case 'p':
              // AM/PM indicator in the locale.  Used in conjunction with �%I�
              // and *not* with �%H�.  An empty string in some locales.
              builder.appendText(ChronoField.AMPM_OF_DAY);
              break;
            case 'O':
              if(i+1>=format.length()) {
                builder.appendLiteral("%O");
              } else {
                switch(format.charAt(++i)) {
                case 'S':
                  int n = 3;
                  if(i+1 < format.length()) {
                      n = Integer.parseInt(Character.toString(format.charAt(++i)));
                  }
                  
                  builder.appendValue(ChronoField.SECOND_OF_MINUTE);
                  if(n > 0) {
                      builder.appendLiteral('.');
                      builder.appendFraction(ChronoField.SECOND_OF_MINUTE,2,2, true);
                  }
                  
                  break;
                default:
                  throw new UnsupportedOperationException("%O[dHImMUVwWy] not yet implemented");
                }
              }
              break;
            case 'S':
              // Second as decimal number (00-61), allowing for up to two
              // leap-seconds (but POSIX-compliant implementations will ignore
              // leap seconds).
              // TODO: I have no idea what the docs are talking about in relation
              // to leap seconds
              builder.appendValue(ChronoField.SECOND_OF_MINUTE,2);
              break;
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

            case 'y':
              // Year without century (00-99). Values 00 to 68 are prefixed by
              // 20 and 69 to 99 by 19 - that is the behaviour specified by
              // the 2004 POSIX standard, but it does also say �it is expected
              // that in a future version the default century inferred from a
              // 2-digit year will change�.
              builder.appendValueReduced(ChronoField.YEAR, 2, 2, 1969);
              break;
            case 'Y':
              // Year with century
              builder.appendValue(ChronoField.YEAR,4);
              break;
            case 'z':
              // Signed offset in hours and minutes from UTC, so �-0800� is 8
              // hours behind UTC.
                builder.appendOffset("+HH:mm", "+0000");
              break;
            case 'Z':
              // (output only.) Time zone as a character string (empty if not
              // available).
              builder.appendZoneOrOffsetId();
              break;
            default:
              throw new UnsupportedOperationException("%" + specifier + " not yet implemented");
            }
          } else {
            builder.appendLiteral(format.substring(i,i+1));
          }
        }
        java.time.format.DateTimeFormatter formatter = builder.toFormatter();
        cache.put(format, formatter);
        return formatter;
    }

    public String format(Date d, String format) {
        StringBuffer sb = new StringBuffer();
        getFormatter(format).formatTo(LocalDateTime.ofInstant(d.toInstant(), ZoneId.systemDefault()), sb);
        return sb.toString();
    }

    public Date parse(String s, String format) {
        java.time.format.DateTimeFormatter formatter = getFormatter(format);
        LocalDateTime local = LocalDateTime.parse(s, formatter);
        ZoneOffset defaultOffset = OffsetDateTime.now(ZoneId.systemDefault()).getOffset();
        return new Date(local.toInstant(defaultOffset).toEpochMilli());
    }

}

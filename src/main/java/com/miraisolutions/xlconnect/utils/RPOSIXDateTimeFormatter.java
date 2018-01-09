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

import java.util.Date;
import java.util.HashMap;
import java.util.Map;
import org.joda.time.format.DateTimeFormatterBuilder;

/**
 * Inspired by http://code.google.com/p/renjin/source/browse/trunk/core/src/main/java/r/base/Time.java?spec=svn379&r=379
 * @author Martin Studer, Mirai Solutions GmbH
 */
public class RPOSIXDateTimeFormatter implements DateTimeFormatter {

    private Map<String, org.joda.time.format.DateTimeFormatter> cache = new HashMap<String, org.joda.time.format.DateTimeFormatter>();

    private org.joda.time.format.DateTimeFormatter getFormatter(String format) {
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
              builder.appendDayOfWeekShortText();
              break;
            case 'A':
              // Full weekday name in the current locale.  (Also matches
              // abbreviated name on input.)
              builder.appendDayOfWeekText();
              break;
            case 'b':
              // Abbreviated month name in the current locale. (Also matches
              // full name on input.)
              builder.appendMonthOfYearShortText();
              break;
            case 'B':
              // Full month name in the current locale.  (Also matches
              // abbreviated name on input.)
              builder.appendMonthOfYearText();
              break;
            case 'c':
              //  Date and time.  Locale-specific on output, �"%a %b %e
              // %H:%M:%S %Y"� on input.
              throw new UnsupportedOperationException("%c not yet implemented");
            case 'd':
              // Day of the month as decimal number (01-31).
              builder.appendDayOfMonth(2);
              break;
            case 'H':
              // Hours as decimal number (00-23).
              builder.appendHourOfDay(2);
              break;
            case 'I':
              // Hours as decimal number (01-12).
              builder.appendHourOfHalfday(2);
              break;
            case 'j':
              // Day of year as decimal number (001-366).
              builder.appendDayOfYear(3);
              break;
            case 'm':
              // Month as decimal number (01-12).
              builder.appendMonthOfYear(2);
              break;
            case 'M':
              // Minute as decimal number (00-59).
              builder.appendMinuteOfHour(2);
              break;
            case 'p':
              // AM/PM indicator in the locale.  Used in conjunction with �%I�
              // and *not* with �%H�.  An empty string in some locales.
              builder.appendHalfdayOfDayText();
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
                  
                  builder.appendSecondOfMinute(2);
                  if(n > 0) {
                      builder.appendLiteral('.');
                      builder.appendFractionOfSecond(n, n);
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
              builder.appendSecondOfMinute(2);
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
              builder.appendTwoDigitYear(1968, true);
              break;
            case 'Y':
              // Year with century
              builder.appendYear(1,4);
              break;
            case 'z':
              // Signed offset in hours and minutes from UTC, so �-0800� is 8
              // hours behind UTC.
              builder.appendTimeZoneOffset(null /* always show offset, even when zero */,
                  true /* show seperators */,
                  1 /* min fields (hour, minute, etc) */,
                  2 /* max fields */ );
              break;
            case 'Z':
              // (output only.) Time zone as a character string (empty if not
              // available).
              builder.appendTimeZoneName();
              break;
            default:
              throw new UnsupportedOperationException("%" + specifier + " not yet implemented");
            }
          } else {
            builder.appendLiteral(format.substring(i,i+1));
          }
        }
        org.joda.time.format.DateTimeFormatter formatter = builder.toFormatter();
        cache.put(format, formatter);
        return formatter;
    }

    public String format(Date d, String format) {
        StringBuffer sb = new StringBuffer();
        getFormatter(format).printTo(sb, new org.joda.time.DateTime(d));
        return sb.toString();
    }

    public Date parse(String s, String format) {
        return getFormatter(format).parseDateTime(s).toDate();
    }

}

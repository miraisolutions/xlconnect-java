/*
 *
    XLConnect
    Copyright (C) 2010 Mirai Solutions GmbH

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

import com.miraisolutions.xlconnect.Workbook;
import java.util.logging.Level;
import java.util.logging.LogManager;
import java.util.logging.Logger;

/**
 *
 * @author Martin Studer, Mirai Solutions GmbH
 */
public abstract class Logging {

    static {
        try {
            LogManager.getLogManager().readConfiguration(Workbook.class.getResourceAsStream("logging.properties"));
        }
        catch(Exception e) {
            e.printStackTrace();
        }
    }

    public static void withLevel(Level logLevel) {
        Logger.getLogger("com.miraisolutions.xlconnect").setLevel(logLevel);
    }

    public static void withLevel(String level) {
        withLevel(Level.parse(level));
    }
}

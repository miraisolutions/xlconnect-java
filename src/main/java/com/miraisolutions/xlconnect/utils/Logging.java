/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
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

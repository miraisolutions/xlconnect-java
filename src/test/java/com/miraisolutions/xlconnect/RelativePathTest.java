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

package com.miraisolutions.xlconnect;

import org.junit.Rule;
import org.junit.Test;
import org.junit.rules.TemporaryFolder;

import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.File;
import java.nio.charset.StandardCharsets;

import static org.junit.Assert.assertEquals;

/**
 * R can change the working directory after rJava has started the JVM. Relative paths must then
 * resolve against the new working directory, not against the 'user.dir' system property, which
 * keeps the directory the JVM started in.
 * <p>
 * Java cannot change its own working directory, so the test starts a second JVM in one directory
 * and points 'user.dir' at a different one.
 */
public class RelativePathTest {

    @Rule
    public TemporaryFolder tmp = new TemporaryFolder();

    @Test
    public void relativePathsUseTheWorkingDirectory() throws Exception {
        File workDir = tmp.newFolder("work");
        File jvmStartDir = tmp.newFolder("jvmStart");

        for (String name : new String[]{"wb.xls", "wb.xlsx"}) {
            Workbook wb = Workbook.getWorkbook(new File(workDir, name), true);
            wb.createSheet("s");
            wb.createName("img", "s!$A$1:$C$5", true, null);
            wb.save();
        }
        ImageIO.write(new BufferedImage(2, 2, BufferedImage.TYPE_INT_RGB), "png", new File(workDir, "img.png"));

        Process p = new ProcessBuilder(
                new File(System.getProperty("java.home"), "bin/java").getPath(),
                "-Duser.dir=" + jvmStartDir.getAbsolutePath(),
                "-cp", System.getProperty("java.class.path"),
                Probe.class.getName(), "wb.xls", "wb.xlsx")
                .directory(workDir)
                .redirectErrorStream(true)
                .start();
        String output = new String(p.getInputStream().readAllBytes(), StandardCharsets.UTF_8);

        assertEquals(output, 0, p.waitFor());
    }

    public static class Probe {
        public static void main(String[] files) throws Exception {
            for (String file : files) {
                Workbook wb = Workbook.getWorkbook(file, false);
                wb.addImage("img.png", "img", true, null);
                wb.save();
            }
        }
    }
}

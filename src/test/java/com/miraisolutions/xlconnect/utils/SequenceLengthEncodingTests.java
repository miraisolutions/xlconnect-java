/*
 *
    XLConnect
    Copyright (C) 2010-2024 Mirai Solutions GmbH

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

import org.junit.Test;

import java.util.Iterator;

import static org.junit.Assert.assertArrayEquals;

public final class SequenceLengthEncodingTests {

    @Test
    public void testEncodingIncrement1() {
        int[] values = {1, 12, 14};
        int[] lengths = {7, 1, 4};
        int increment = 1;
        int[] expected = {1, 2, 3, 4, 5, 6, 7, 12, 14, 15, 16, 17};
        SequenceLengthEncoding sle = new SequenceLengthEncoding(values, lengths, increment);

        assertArrayEquals(sleToIntArray(sle), expected);
    }

    private SequenceLengthEncoding createSleIncrement2() {
        int[] values = {2, 9, 12};
        int[] lengths = {4, 1, 3};
        int increment = 2;
        return new SequenceLengthEncoding(values, lengths, increment);
    }

    @Test
    public void testEncodingIncrement2() {
        SequenceLengthEncoding sle = createSleIncrement2();
        int[] expected = {2, 4, 6, 8, 9, 12, 14, 16};
        assertArrayEquals(sleToIntArray(sle), expected);
    }

    @Test
    public void testRepeatingEncodingIncrement2() {
        SequenceLengthEncoding sle = createSleIncrement2();
        int[] expected = {2, 4, 6, 8, 9, 12, 14, 16, 2, 4, 6, 8, 9, 12, 14, 16, 2, 4, 6};
        assertArrayEquals(sleToIntArray(sle, expected.length), expected);
    }

    private static int[] sleToIntArray(SequenceLengthEncoding sle, int collect) {
        boolean repeating = collect > 0;
        collect = repeating ? collect : sle.length();
        int[] values = new int[collect];
        Iterator<Integer> it = sle.iterator(repeating);
        for (int i = 0; i < collect; i++) {
            values[i] = it.next();
        }
        return values;
    }

    private static int[] sleToIntArray(SequenceLengthEncoding sle) {
        return sleToIntArray(sle, 0);
    }
}

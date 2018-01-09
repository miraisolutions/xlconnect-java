/*
 *
    XLConnect
    Copyright (C) 2017-2018 Mirai Solutions GmbH

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

import java.util.Iterator;

/**
 * Sequence Length Encoding
 *
 * Encodes a sequence of values as a set of sub-sequences with a certain step size (increment).
 */
public class SequenceLengthEncoding implements RepeatableIterable<Integer> {
    // Start values of sub-sequences
    private final int[] values;
    // Sub-sequence lengths
    private final int[] cumLengths;
    // Sequence increment / step size
    private final int increment;

    public SequenceLengthEncoding(int[] values, int[] lengths, int increment) {
        if(values.length != lengths.length) throw new IllegalArgumentException("Arrays must be of same length");
        for(int i = 0; i < lengths.length; i++) {
            if(lengths[i] < 1) throw new IllegalArgumentException("Lengths must be greater than zero!");
        }

        this.values = values;
        this.cumLengths = cumulativeLengths(lengths);
        this.increment = increment;
    }

    /**
     * Creates an iterator which iterates over this sequence.
     * @param repeating if `true`, the iterator continues to iterate over this sequence (i.e. it 'resets' and loops
     * over this sequence; `hasNext` always returns `true`); if `false`, the iterator stops when reaching the end
     * of this sequence
     * @return Sequence iterator
     */
    public Iterator<Integer> iterator(boolean repeating) {
        return new SequenceIterator(repeating);
    }

    /** Creates an non-repeating sequence iterator */
    public Iterator<Integer> iterator() {
        return iterator(false);
    }

    /** Sequence length */
    public int length() {
        return cumLengths[cumLengths.length - 1];
    }

    /** Calculates the cumulative length of the sub-sequences */
    private static int[] cumulativeLengths(int[] lengths) {
        int[] cum = new int[lengths.length];
        int total = 0;
        for(int i = 0; i < lengths.length; i++) {
            total += lengths[i];
            cum[i] = total;
        }
        return cum;
    }

    private class SequenceIterator implements Iterator<Integer> {
        // Do we repeat iterating?
        private boolean repeating = false;
        private int i = 0;
        private int chunk = 0;

        public SequenceIterator(boolean repeating) {
            this.repeating = repeating;
        }

        // Number of elements in previous chunks
        private int elemsInPrevChunks() {
            int elems;
            if(chunk == 0) {
                elems = 0;
            } else {
                elems = cumLengths[chunk - 1];
            }
            return elems;
        }

        public boolean hasNext() {
            return repeating || i < length();
        }

        public Integer next() {
            int result = values[chunk] + (i - elemsInPrevChunks()) * increment;
            i += 1;
            if(repeating && i >= SequenceLengthEncoding.this.length()) {
                i = 0;
                chunk = 0;
            } else if(i >= cumLengths[chunk]) {
                chunk += 1;
            }
            return result;
        }

        public void remove() {}
    }
}

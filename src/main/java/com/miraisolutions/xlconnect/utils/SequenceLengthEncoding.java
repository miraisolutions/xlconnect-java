package com.miraisolutions.xlconnect.utils;

import java.util.Iterator;

public class SequenceLengthEncoding {
    private final int[] values;
    private final int[] cumLengths;

    public SequenceLengthEncoding(int[] values, int[] lengths) {
        if(values.length != lengths.length) throw new IllegalArgumentException("Arrays must be of same length");
        for(int i = 0; i < lengths.length; i++) {
            if(lengths[i] < 1) throw new IllegalArgumentException("Lengths must be greater than zero!");
        }

        this.values = values;
        this.cumLengths = cumulativeLengths(lengths);
    }

    public Iterator<Integer> iterator() {
        return new SequenceIterator();
    }

    public int length() {
        return cumLengths[cumLengths.length - 1];
    }

    private static int[] cumulativeLengths(int[] lengths) {
        int[] cum = new int[lengths.length];
        int total = 0;
        for(int i = 0; i < lengths.length; i++) {
            total += lengths[i];
            cum[i] = total;
        }
        return cum;
    }

    public class SequenceIterator implements Iterator<Integer> {
        private int i = 0;
        private int chunk = 0;

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

        @Override
        public boolean hasNext() {
            return i < length();
        }

        @Override
        public Integer next() {
            int result = values[chunk] + (i - elemsInPrevChunks());
            i += 1;
            if(i >= cumLengths[chunk]) chunk += 1;
            return result;
        }

        @Override
        public void remove() {}
    }
}

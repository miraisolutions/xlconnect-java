package com.miraisolutions.xlconnect.utils;

import java.util.Iterator;

/**
 * A repeatable iterable sequence wrapper around an array of elements
 * @param <T> Element type
 */
public class SimpleSequence<T> implements RepeatableIterable<T> {
    private T[] values;

    public static SimpleSequence<String> create(String[] values) {
        return new SimpleSequence<String>(values);
    }

    public static SimpleSequence<Integer> create(int[] values) {
        Integer[] newValues = new Integer[values.length];
        int i = 0;
        for (int value : values) {
            newValues[i++] = Integer.valueOf(value);
        }
        return new SimpleSequence<Integer>(newValues);
    }

    public static SimpleSequence<Double> create(double[] values) {
        Double[] newValues = new Double[values.length];
        int i = 0;
        for (double value : values) {
            newValues[i++] = Double.valueOf(value);
        }
        return new SimpleSequence<Double>(newValues);
    }

    public SimpleSequence(T[] values) {
        this.values = values;
    }

    public Iterator<T> iterator(boolean repeating) {
        return new SequenceIterator(repeating);
    }

    public int length() {
        return values.length;
    }

    public Iterator<T> iterator() {
        return iterator(false);
    }

    private class SequenceIterator implements Iterator<T> {
        private boolean repeating = false;
        private int i = 0;

        public SequenceIterator(boolean repeating) {
            this.repeating = repeating;
        }

        public boolean hasNext() {
            return repeating || i < values.length;
        }

        public T next() {
            T result = values[i];
            i += 1;
            if(repeating && i >= length()) i = 0;
            return result;
        }

        public void remove() {}
    }
}

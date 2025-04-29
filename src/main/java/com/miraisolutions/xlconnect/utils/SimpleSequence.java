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

import java.util.Arrays;
import java.util.Iterator;
import java.util.NoSuchElementException;

/**
 * A repeatable iterable sequence wrapper around an array of elements
 *
 * @param <T> Element type
 */
public final class SimpleSequence<T> implements RepeatableIterable<T> {
    private final T[] values;

    public static SimpleSequence<String> create(String[] values) {
        return new SimpleSequence<>(values);
    }

    public static SimpleSequence<Integer> create(int[] values) {
        Integer[] newValues = Arrays.stream(values).boxed().toArray(Integer[]::new);
        return new SimpleSequence<>(newValues);
    }

    public static SimpleSequence<Double> create(double[] values) {
        Double[] newValues = Arrays.stream(values).boxed().toArray(Double[]::new);
        return new SimpleSequence<>(newValues);
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

    private final class SequenceIterator implements Iterator<T> {
        private final boolean repeating;
        private int i = 0;

        public SequenceIterator(boolean repeating) {
            this.repeating = repeating;
        }

        public boolean hasNext() {
            return repeating || i < values.length;
        }

        public T next() {
            if (!hasNext()) {
                throw new NoSuchElementException();
            }

            T result = values[i % values.length];
            i++;
            return result;
        }
    }
}

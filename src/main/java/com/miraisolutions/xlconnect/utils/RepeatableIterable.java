package com.miraisolutions.xlconnect.utils;

import java.util.Iterator;

/**
 * Iterable that supports repeating/replicating iterators
 * @param <T> Element type
 */
public interface RepeatableIterable<T> extends Iterable<T> {
    Iterator<T> iterator(boolean repeating);
    int length();
}

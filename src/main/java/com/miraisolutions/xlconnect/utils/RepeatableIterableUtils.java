package com.miraisolutions.xlconnect.utils;

import java.util.Arrays;
import java.util.Collections;
import java.util.Iterator;

/**
 * Repeatable iterable utilities
 */
public class RepeatableIterableUtils {
    /** Function with four type parameters */
    public interface Function4<T1, T2, T3, T4> {
        void apply(T1 t1, T2 t2, T3 t3, T4 t4);
    }

    /** Function with five type parameters */
    public interface Function5<T1, T2, T3, T4, T5> {
        void apply(T1 t1, T2 t2, T3 t3, T4 t4, T5 t5);
    }

    /**
     * Apply a `Function4` across the specified repeatable iterables where elements
     * are replicated to the length of the longest iterable.
     */
    public static <T1, T2, T3, T4> void foreach(RepeatableIterable<T1> r1, RepeatableIterable<T2> r2,
                                                RepeatableIterable<T3> r3, RepeatableIterable<T4> r4,
                                                Function4 function) {

        int maxLen = getMaxLength(new RepeatableIterable[]{r1, r2, r3, r4});
        Iterator<T1> i1 = r1.iterator(true);
        Iterator<T2> i2 = r2.iterator(true);
        Iterator<T3> i3 = r3.iterator(true);
        Iterator<T4> i4 = r4.iterator(true);
        for (int i = 0; i < maxLen; i++) function.apply(i1.next(), i2.next(), i3.next(), i4.next());
    }

    /**
     * Apply a `Function5` across the specified repeatable iterables where elements
     * are replicated to the length of the longest iterable.
     */
    public static <T1, T2, T3, T4, T5> void foreach(RepeatableIterable<T1> r1, RepeatableIterable<T2> r2,
                                                    RepeatableIterable<T3> r3, RepeatableIterable<T4> r4,
                                                    RepeatableIterable<T5> r5, Function5 function) {

        int maxLen = getMaxLength(new RepeatableIterable[]{r1, r2, r3, r4, r5});
        Iterator<T1> i1 = r1.iterator(true);
        Iterator<T2> i2 = r2.iterator(true);
        Iterator<T3> i3 = r3.iterator(true);
        Iterator<T4> i4 = r4.iterator(true);
        Iterator<T5> i5 = r5.iterator(true);
        for (int i = 0; i < maxLen; i++) function.apply(i1.next(), i2.next(), i3.next(), i4.next(), i5.next());
    }

    private static int getMaxLength(RepeatableIterable<?> it[]) {
        Integer[] lengths = new Integer[it.length];
        for (int i = 0; i < it.length; i++) lengths[i] = it[i].length();
        return Collections.max(Arrays.asList(lengths));
    }
}

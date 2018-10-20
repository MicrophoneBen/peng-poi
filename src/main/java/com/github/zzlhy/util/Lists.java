package com.github.zzlhy.util;

import java.util.ArrayList;
import java.util.Collections;

/**
 * Created by Administrator on 2018-10-20.
 */
public class Lists {
    public static <E> ArrayList<E> newArrayList(E... elements) {
        if(elements == null){
            throw new NullPointerException();
        }
        ArrayList<E> list = new ArrayList();
        Collections.addAll(list, elements);
        return list;
    }
}

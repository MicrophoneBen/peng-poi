package com.github.zzlhy.func;

/**
 * 保存数据的方法
 * Created by Administrator on 2018-10-16.
 */
@FunctionalInterface
public interface SaveDataHandler<T> {

    void save(T t);
}

package com.github.zzlhy.func;

/**
 * 已处理数变化时的回调
 * Created by Administrator on 2018-10-16.
 */
@FunctionalInterface
public interface IndexChangeHandler {

    /**
     * 已处理数变化时的回调
     * @param index 已处理条数
     */
    void change(int index);
}

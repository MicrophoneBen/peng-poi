package com.github.zzlhy.func;

/**
 * 对数据对象进行验证是否合格
 * Created by Administrator on 2018-10-16.
 */
@FunctionalInterface
public interface ValidateDataHandler<T> {

    /**
     * 数据验证回调
     * @param t 对象
     * @return 验证不通过时的原因,验证通过时,返回null或空字符串
     */
    String valid(T t);
}

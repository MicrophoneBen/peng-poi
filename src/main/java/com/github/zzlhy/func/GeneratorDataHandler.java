package com.github.zzlhy.func;


import java.util.List;

/**
 *
 * Created by Administrator on 2018-10-12.
 */
@FunctionalInterface
public interface GeneratorDataHandler {

    /**
     * 生成数据的方法
     * @param page 当前页码
     * @param pageSize 每页条数
     * @return 数据集合
     */
    List<?> generatorData(int page, int pageSize);
}

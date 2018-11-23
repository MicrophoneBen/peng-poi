package com.github.zzlhy.main;

import com.github.zzlhy.entity.TableParam;
import com.github.zzlhy.func.IndexChangeHandler;
import com.github.zzlhy.func.SaveDataHandler;
import com.github.zzlhy.func.ValidateDataHandler;

import java.io.InputStream;
import java.util.List;
import java.util.Map;

/**
 * Excel导入接口
 * Created by zpeng on 2018-11-15.
 */
public interface ImportAbstract {

    /**
     * 读取Excel (适用于少量数据,一次性读取所以Excel里的数据)
     * @param stream 流
     * @param tableParam Excel配置对象
     * @param clazz Class
     * @return 返回读取到的数据集合
     */
    List<?> readExcel(InputStream stream, TableParam tableParam, Class<?> clazz);


    /**
     * 读取Excel数据并执行方法
     * 说明：支持大量数据导入,减少内存消耗,支持多个sheet读取,支持xls和xlsx
     * @param stream 流
     * @param tableParam Excel配置对象
     * @param clazz Class
     * @param saveDataHandler 保存数据的方法
     * @param validateDataHandler 效验数据的方法
     * @param indexChangeHandler 处理条数变化的回调
     * @return 导入失败的数据
     */
    List<Map<String,Object>> readAndExecute(InputStream stream, TableParam tableParam, Class<?> clazz, SaveDataHandler saveDataHandler, ValidateDataHandler validateDataHandler, IndexChangeHandler indexChangeHandler);
}

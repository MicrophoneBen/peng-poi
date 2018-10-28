import com.github.zzlhy.ExcelUtils;
import com.github.zzlhy.entity.Col;
import com.github.zzlhy.entity.TableParam;
import com.github.zzlhy.util.Lists;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.Test;

import java.beans.IntrospectionException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.reflect.InvocationTargetException;
import java.text.ParseException;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

/**
 * Created by Administrator on 2018-10-20.
 */
public class TestPOI {

    @Test
    public void test1() throws IllegalAccessException, IntrospectionException, InvocationTargetException, IOException {
        //导出参数配置
        List<Col> cols = Lists.newArrayList(
                new Col("姓名","name"),
                new Col("年龄","age"),
                new Col("生日","birthday"),
                new Col("启用","active"),
                new Col("创建日期","createDate",25),
                new Col("统计",30,"B2:B1000+1"),
                new Col("随机数",30,"int(RAND()*1000)")
        );
        TableParam tableParam=new TableParam(cols);
        List<Person> list = new ArrayList<>();
        list.add(new Person("小泽",30, LocalDate.of(1988,10,2),true, LocalDateTime.now()));
        list.add(new Person("小苍",31, LocalDate.of(1987,1,2),true, LocalDateTime.now()));
        list.add(new Person("吉泽",28, LocalDate.of(1988,10,  2),true, LocalDateTime.now()));
        list.add(new Person("武藤",20, LocalDate.of(1980,12,2),true, LocalDateTime.now()));
        list.add(new Person("波多",18, LocalDate.of(1990,8,15),true, LocalDateTime.now()));

        tableParam.setFreezeColSplit(1);
        tableParam.setFreezeRowSplit(1);
        Workbook workbook = ExcelUtils.excelExport(tableParam,list);

        String name = "D:\\"+LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyyMMddHHmmss"))+".xlsx";
        workbook.write(new FileOutputStream(name));
        System.out.println("OK,"+name);
    }

    @Test
    public void test2() throws IllegalAccessException, IntrospectionException, InvocationTargetException, IOException, ParseException, InvalidFormatException, InstantiationException {
        //导出参数配置
        List<Col> cols = Lists.newArrayList(
                new Col("姓名","name"),
                new Col("年龄","age"),
                new Col("生日","birthday"),
                new Col("启用","active"),
                new Col("创建日期","createDate",25)
        );
        TableParam tableParam=new TableParam(cols);
        tableParam.setFreezeColSplit(1);
        tableParam.setFreezeRowSplit(1);
        List<Person> list = (List<Person>) ExcelUtils.readExcel("D:\\20181028154845.xlsx", tableParam, Person.class);

        System.out.println("OK");
    }

    @Test
    public void test3(){
        System.out.println(Math.round(2.00001));
    }
}

import com.github.zzlhy.entity.*;
import com.github.zzlhy.entity.Color;
import com.github.zzlhy.main.ExcelExport;
import com.github.zzlhy.main.ExportStyleImpl;
import com.github.zzlhy.util.Lists;
import com.github.zzlhy.util.Utils;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;
import org.junit.Test;

import java.beans.IntrospectionException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.reflect.InvocationTargetException;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

/**
 * Created by Administrator on 2018-10-20.
 */
public class TestPOI {

    //@Test
    //public void test1() throws IllegalAccessException, IntrospectionException, InvocationTargetException, IOException {
    //    //导出参数配置
    //    List<Col> cols = Lists.newArrayList(
    //            new Col("姓名","name",ColStyle.of().setHorizontalAlignment(HorizontalAlignment.CENTER).setVerticalAlignment(VerticalAlignment.CENTER)),
    //            Col.of("年龄","age",new String[]{"10","20","30"}),
    //            new Col("生日","birthday",30,ColStyle.of(Color.ORANGE.getIndex())),
    //            new Col("启用","active",ColStyle.of().setHorizontalAlignment(HorizontalAlignment.RIGHT)),
    //            new Col("创建日期","createDate",25,ColStyle.of(FontStyle.of(Color.RED.getIndex()).setUnderline((byte)1))),
    //            new Col("统计",30,"B2:B1000+1").setDropdownList(new String[]{"A","B","C"}),
    //            new Col("随机数",10,"int(RAND()*1000)")
    //    );
    //
    //    List<Person> list = new ArrayList<>();
    //    list.add(new Person("小泽",30, Date.from(LocalDate.of(1988,10,2).atStartOfDay(ZoneId.systemDefault()).toInstant()),true, LocalDateTime.now()));
    //    list.add(new Person("小苍",31, Date.from(LocalDate.of(1987,1,2).atStartOfDay(ZoneId.systemDefault()).toInstant()),true, LocalDateTime.now()));
    //    list.add(new Person("吉泽",28, Date.from(LocalDate.of(1988,10,  2).atStartOfDay(ZoneId.systemDefault()).toInstant()),true, LocalDateTime.now()));
    //    list.add(new Person("武藤",20, Date.from(LocalDate.of(1980,12,2).atStartOfDay(ZoneId.systemDefault()).toInstant()),true, LocalDateTime.now()));
    //    list.add(new Person("波多",18, Date.from(LocalDate.of(1990,8,15).atStartOfDay(ZoneId.systemDefault()).toInstant()),true, LocalDateTime.now()));
    //
    //    Workbook workbook = ExcelExport.exportExcelByObject(
    //            TableParam.of(cols),
    //            list,new ExportStyleImpl()
    //    );
    //
    //    String name = "D:\\"+LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyyMMddHHmmss"))+".xlsx";
    //    workbook.write(new FileOutputStream(name));
    //    System.out.println("OK,"+name);
    //
    //}

    //@Test
    //public void test2() throws IllegalAccessException, IntrospectionException, InvocationTargetException, IOException, ParseException, InvalidFormatException, InstantiationException {
    //    //导出参数配置
    //    List<Col> cols = Lists.newArrayList(
    //            new Col("姓名","name"),
    //            new Col("年龄","age"),
    //            new Col("生日","birthday","yyyy-MM-dd"),
    //            new Col("启用","active"),
    //            new Col("创建日期","createDate",25)
    //    );
    //    List<Person> list = (List<Person>) ExcelUtils.readExcel("D:\\20181106184143.xlsx", TableParam.of(cols), Person.class);
    //
    //    System.out.println("OK");
    //}
    //
    //@Test
    //public void test3(){
    //    System.out.println(Math.round(2.00001));
    //}

    //自定义样式测试
    @Test
    public void testStyle() throws IllegalAccessException, IntrospectionException, InvocationTargetException, IOException {
        //导出参数配置
        List<Col> cols = Lists.newArrayList(
                new Col("姓名","name",ColStyle.of().setHorizontalAlignment(HorizontalAlignment.CENTER).setVerticalAlignment(VerticalAlignment.CENTER).setBackgroundColor(Color.LIGHT_ORANGE.getIndex())),
                Col.of("年龄","age",new String[]{"10","20","30"}),
                new Col("生日","birthday",30,ColStyle.of(Color.ORANGE.getIndex())),
                new Col("启用","active",ColStyle.of().setHorizontalAlignment(HorizontalAlignment.RIGHT).setTopBorder(BorderStyle.DASHED).setLeftBorder(BorderStyle.DOUBLE).setLeftBorderColor(Color.GREEN.getIndex())),
                new Col("创建日期","createDate",25,ColStyle.of(FontStyle.of(Color.RED.getIndex()).setUnderline((byte)1))),
                new Col("统计",30,"B2:B1000+1").setDropdownList(new String[]{"A","B","C"}),
                new Col("随机数",10,"int(RAND()*1000)")
        );

        List<Person> list = new ArrayList<>();
        list.add(new Person("小泽",30, Date.from(LocalDate.of(1988,10,2).atStartOfDay(ZoneId.systemDefault()).toInstant()),true, LocalDateTime.now()));
        list.add(new Person("小苍",31, Date.from(LocalDate.of(1987,1,2).atStartOfDay(ZoneId.systemDefault()).toInstant()),true, LocalDateTime.now()));
        list.add(new Person("吉泽",28, Date.from(LocalDate.of(1988,10,  2).atStartOfDay(ZoneId.systemDefault()).toInstant()),true, LocalDateTime.now()));
        list.add(new Person("武藤",20, Date.from(LocalDate.of(1980,12,2).atStartOfDay(ZoneId.systemDefault()).toInstant()),true, LocalDateTime.now()));
        list.add(new Person("波多",18, Date.from(LocalDate.of(1990,8,15).atStartOfDay(ZoneId.systemDefault()).toInstant()),true, LocalDateTime.now()));

        Workbook workbook = ExcelExport.exportExcelByObject(
                TableParam.of(cols),
                list,new ExportStyleImpl()
        );

        String name = "D:\\"+LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyyMMddHHmmss"))+".xlsx";
        workbook.write(new FileOutputStream(name));
        System.out.println("OK,"+name);

    }
}

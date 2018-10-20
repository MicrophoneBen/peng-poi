# peng-poi

基于POI封装的支持大数据量的Excel导入导出库,切合实际,
无额外无用功能,简单而实用.经某大型进销存管理系统实践检验.

POI based Excel import and export library which supports large amounts of data is practical.

No extra useless function is simple and practical. It is tested by a large inventory management system.

**Quick Start**

Add Maven Dependency:

```
<dependency>
    <groupId>com.github.zzlhy</groupId>
    <artifactId>peng-poi</artifactId>
    <version>1.0</version>
</dependency>

```
Or Add Gradle:

```
compile group: 'com.github.zzlhy', name: 'peng-poi', version: '1.0'
```

**Usage:**
##### export example1(最简单的导出方式):
```
    //数据列表
    List<User> list = getListUser();
    //列配置
    List<Col> cols = Lists.newArrayList(
        new Col("用户名","userName"),
        new Col("姓名","name"),
        new Col("手机号码","telephone"),
        new Col("邮箱","email")
    );
    TableParam tableParam=new TableParam(cols);
    Workbook workbook = ExcelUtils.excelExportByObject(tableParam,list);

```
##### export example2(最简单的导出方式):
```
    //数据列表,数据集合里面是Map
    List<Map<String,Object>> list = getMapList();
    //列配置
    List<Col> cols = Lists.newArrayList(
            new Col("用户名","userName"),
            new Col("姓名","name"),
            new Col("手机号码","telephone"),
            new Col("邮箱","email")
    );
    TableParam tableParam=new TableParam(cols);
    Workbook workbook = ExcelUtils.excelExportByMap(tableParam,list);

```

##### export example3(大量数据导出):
```
    //分页方式
    //要导出的数据总条数
    long toltal = getDataTotal();
    //列配置
    List<Col> cols = Lists.newArrayList(
            new Col("用户名","userName"),
            new Col("姓名","name"),
            new Col("手机号码","telephone"),
            new Col("邮箱","email")
        );
    TableParam tableParam=new TableParam(cols);
    SXSSFWorkbook workbook = ExcelUtils.excelExport(tableParam,total,(page,size)-> userService.list(page, size));

    说明：(page,size)-> userService.list(page, size) page是当前页码,从0开始,size是每页大小,
    userService.list(page, size)这个方法就是分页查询获取数据的方法
```
##### import example1(Excel读取):
```
    //列配置
    List<Col> cols = Lists.newArrayList(
            new Col("userName"),
            new Col("name"),
            new Col("telephone"),
            new Col("email")
        );
    TableParam tableParam=new TableParam(cols);
    List<User> list = ExcelUtils.readExcel(inputStream,tableParam,SysUser.class);
```

##### import example2(Excel读取导入):
```
    //说明：读取到数据后可以直接加入数据库
    //列配置
    List<Col> cols = Lists.newArrayList(
                new Col("userName"),
                new Col("name"),
                new Col("telephone"),
                new Col("email")
            );
    TableParam tableParam=new TableParam(cols);
    List<Map<String,Object>> failData= ExcelUtils.excelImport(inputStream,tableParam,SysUser.class, 
        user -> {
            //保存用户的方法
            userService.saveOrUpdate(user);
        },
        user -> validate((SysUser)user), //validate()为数据验证的方法,比如判断唯一性,非空,长度等,需要返回错误信息,会在失败数据中体现
        index -> System.out.println(index) //执行条数变化的回调方法
     );
     注:返回的是导入失败的数据
```
#### TableParam对象参数:
```
    //Sheet名称
    String sheetName="Sheet";
    //每个sheet允许的数据行数,超过此行数会创建新的sheet (xlsx时最大值为1048576)
    int sheetDataTotal = 500000;
    //导出时开始写入的起始行,默认从0开始
    int startRow=0;
    //导入时,起始读取行
    int readRow=1;
    //行高度
    float height = 15;
    //是否创建标题行
    boolean createHeadRow=true;
    //标题行设置
    HeadRowStyle headRowStyle=new HeadRowStyle();
    //导出Excel类型xls/xlsx   默认为xlsx
    ExcelType excelType=ExcelType.XLSX;
    //列属性数组
    List<Col> cols;
```
#### Col对象参数:
```
       //标题名称
       String title;
       //属性名
       String key;
       //列宽度,默认为15个字符
       int width = 15;
       //日期格式
       String format;
       //公式或函数
       String formula;
       //对读取到的单元格值进行转换的实现,可自行实现
       ConvertValue convertValue;
       //列样式
       ColStyle colStyle=new ColStyle();
```

**License:**

This project is licensed under the [Apache 2 License](http://www.apache.org/licenses/LICENSE-2.0).




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
    List<Col> cols = new ArrayList<>();
    cols.add(new Col("用户名","userName"));
    cols.add(new Col("姓名","name"));
    cols.add(new Col("手机号码","telephone"));
    cols.add(new Col("邮箱","email"));
    TableParam tableParam=new TableParam(cols);
    Workbook workbook = ExcelUtils.excelExportByObject(tableParam,list);

```
##### export example2(最简单的导出方式):
```
    //数据列表,数据集合里面是Map
    List<Map<String,Object>> list = getMapList();
    //列配置
    List<Col> cols = new ArrayList<>();
    cols.add(new Col("用户名","userName"));
    cols.add(new Col("姓名","name"));
    cols.add(new Col("手机号码","telephone"));
    cols.add(new Col("邮箱","email"));
    TableParam tableParam=new TableParam(cols);
    Workbook workbook = ExcelUtils.excelExportByMap(tableParam,list);

```

##### export example3(大量数据导出):
```
    //分页方式
    //要导出的数据总条数
    long toltal = getDataTotal();
    //列配置
    List<Col> cols = new ArrayList<>();
    cols.add(new Col("用户名","userName"));
    cols.add(new Col("姓名","name"));
    cols.add(new Col("手机号码","telephone"));
    cols.add(new Col("邮箱","email"));
    TableParam tableParam=new TableParam(cols);
    SXSSFWorkbook workbook = ExcelUtils.excelExport(tableParam,total,(page,size)-> userService.list(page, size));

    说明：(page,size)-> userService.list(page, size) page是当前页码,从0开始,size是每页大小,
    userService.list(page, size)这个方法就是分页查询获取数据的方法
```
##### import example1(Excel读取):
```
    //列配置
    List<Col> cols = new ArrayList<>();
    cols.add(new Col("userName"));
    cols.add(new Col("name"));
    cols.add(new Col("telephone"));
    cols.add(new Col("email"));
    TableParam tableParam=new TableParam(cols);
    List<User> list = ExcelUtils.readExcel(inputStream,tableParam,SysUser.class);
```

##### import example1(Excel读取导入):
```
    //说明：读取到数据后可以直接加入数据库
    //列配置
    List<Col> cols = new ArrayList<>();
    cols.add(new Col("userName"));
    cols.add(new Col("name"));
    cols.add(new Col("telephone"));
    cols.add(new Col("email"));
    TableParam tableParam=new TableParam(cols);
    List<Map<String,Object>> failData= ExcelUtils.excelImport(inputStream,tableParam,SysUser.class, 
        user -> {
            //保存用户的方法
            userService.saveOrUpdate(user);
        },
        user -> validate((SysUser)user), //validate为数据验证的方法,比如判断唯一性,非空,长度等,未验证通过的需要返回错误信息,会在失败数据中体现
        index -> System.out.println(index) //执行条数变化的回调方法
     );
     注:返回的是导入失败的数据
```


**License:**

This project is licensed under the [Apache 2 License](http://www.apache.org/licenses/LICENSE-2.0).




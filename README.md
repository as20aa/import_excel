# 从excel导入数据库
## 说明
程序实现的是：		
* 从excel文件中读取数据
* 将读取的数据存在datatable中
* 将datatable中的数据导入到SQL Server中		
## 使用说明：		
* 找到import_excel.exe 并运行
* 依次输入datasource,userid,password,initcatalog
* 点击 login
* 登陆成功后点击“浏览”选定目标文件
* 点击导入datatable（由于用的com速度比较慢，但是如果是按完这个按钮之后点左边datagridview框可以加速？？？）
* 点击导入SQL Server
* 如果导入的数据有重复记录，则会询问是否需要更新，更新是删除原有的记录再导入新的记录
# bug
* 导入到datatable速度很慢
* 未加入对database和table存在时的操作，如果已存在是追加打开还是直接覆盖，如果是不存在数据库是否要创建
* 数据库连接时无设置使用的database和table选项
* 如果输入的列表中有一行的数据类型出错则程序会整个退出，无表格数据检查
* 无对数据库现有的database或table的选择
* 不能追加第一列数据相同的行数据
* 无应用程序的默认配置和账号记忆功能

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
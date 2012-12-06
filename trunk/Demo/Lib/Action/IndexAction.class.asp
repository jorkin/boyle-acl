<%
'// 本类由系统自动生成，仅供测试用途
With System.Template
	'// 载入页面
	'.setCache = "boyletpl,4,60"
	.Root = .Root & "/Index/"
	.File = "index.html"

	'// 设置数据库表的前缀
	C("DB.PREFIX") = "BE_"

	'// 对模板中的所有普通标签赋值
	.d("$") = System.Data.Command("SELECT CATE_NAME AS CATENAME FROM [BE_CATEGORY] WHERE ID=?", Array(Array("id",3,1,4,2)))

	.d("title") = "欢迎使用Boyle.ACL框架 - 模板示例"

	'// sql="select top 10 id,ci_name as name from be_customer"
	'// 直接在模板中指定SQL语句，将自动执行下面的操作，这样做不安全
	'Dim Rs: Set Rs = System.Data.Read(System.Data.ToSQL(Array("BE_CUSTOMER", "ID,CI_NAME,CI_TELEPHONE,CI_ADDRESS", 10), "", ""))
	'.d("customer") = Array(Rs, "ID,NAME,TEL,ADDRESS")

	Set Customer = M("CUSTOMER")
	Customer.Parameters("") = Array("LIMIT:10", "FIELD:ID,CI_NAME AS NAME,CI_TELEPHONE AS TEL,CI_ADDRESS AS ADDR")
	Customer.Parameters("FIELD") = Array("ID", "CI_NAME", "CI_TELEPHONE", "CI_ADDRESS")
	Customer.Parameters("WHERE") = Array("id>100 and ci_address<>'-1'", "CI_TELEPHONE='-1'", "_logic:OR")
	Customer.Parameters("ORDER") = Array("ID DESC")
	'Customer.Parameters("") = 10 '// 相当于Customer.Parameters("LIMIT") = 10
	'Customer.Parameters("FIELD") = "ID,CI_NAME AS NAME,CI_TELEPHONE,CI_ADDRESS AS ADDRESS"
	.d("customer") = Customer.Select()
	.d("sql") = Customer.Parameters("SQL")

	'With System.Data
	''	.Page.Parameters("") = Array("CURRENTPAGE:"&blPage&"", "PAGESIZE:15")
	''	.Page.SQL = .ToSQL(Array("BE_PARTS", "ID,CP_NAME,CP_LOCALITY,CP_CAR", ""), "", "")
	''	Dim blPageData: blPageData = System.Array.Swap(.Page.Run)
	'End With
	'.d("customerpage") = Array(blPageData, "id,name,locality,car")
	'.d("pager") = System.Data.Page.Out
	'.d("sql") = System.Data.Page.Parameters("SQL")

	Dim blPage: blPage = System.Get("PAGE", 0)
	Dim Parts: Set Parts = M("PARTS")
	Parts.Parameters("") = Array("CURRENTPAGE:"&blPage&"", "PAGESIZE:15", "FIELD:ID,CP_NAME,CP_LOCALITY,CP_CAR")
	Dim PagerResult: PagerResult = Parts.Pager()
	.d("customerpage") = Array(PagerResult(0), "id,name,locality,car")
	.d("pager") = PagerResult(1)
	'.d("sql") = PagerResult(2)("SQL")

	'// 输出页面
	.Display

	System.Data.C(Rs)
End With


'// D函数用于实例化Model 格式 项目://分组/模块
'Public Function D()
'End Function

'// M函数用于实例化一个没有模型文件的Model
'Public Function M()
'End Function

'// 缓存管理
Public Function S()
End Function

'// 快速文件数据读取和保存 针对简单类型数据 字符串、数组
Public Function F()
End Function

'// URL组装 支持不同URL模式
Public Function U()
End Function

'// 获取和设置语言定义(不区分大小写)
Public Function L()
End Function
%>
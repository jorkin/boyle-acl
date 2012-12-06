<%
'// +--------------------------------------------------------------------------
'// | Boyle.ACL [系统数据库操作类]
'// +--------------------------------------------------------------------------
'// | Copyright (c) 2008-2012 By Boyle. All rights reserved.
'// +--------------------------------------------------------------------------
'// | Licensed ( http://www.gnu.org/licenses/gpl.html )
'// +--------------------------------------------------------------------------
'// | Author: Boyle <boyle7[at]qq.com>
'// +--------------------------------------------------------------------------

Class Cls_Data

	'// 声明私有对象
	Private PrPager
	Private m_Conn, m_Source, m_ConnString
	
	'// 声明公共对象

	'// 初始化资源
	Private Sub Class_Initialize()
		System.Error.E(200) = "错误的参数个数！"
		System.Error.E(201) = "数据库服务器端连接错误，请检查数据库连接信息是否正确！"
		System.Error.E(202) = "仅支持从MS SQL Server数据库中调用存储过程！"
		System.Error.E(203) = "生成Json格式代码出错！"
	End Sub
	
	'// 释放资源
	Private Sub Class_Terminate()
		If IsObject(PrPager) Then Set PrPager = Nothing End If
	End Sub
	
	'// 声明模块单元
	Public Property Get Page()
		If Not IsObject(PrPager) Then Set PrPager = New Cls_Data_Page End If
		Set Page = PrPager
	End Property
	
	'// 新建类实例
	Public Function [New]()
		Set [New] = New Cls_Data
	End Function
	
	'// 数据库地址（读/写）
	Public Property Get Source()
		Source = m_Source
	End Property
	Public Property Let Source(ByVal blParam)
		m_Source = blParam
	End Property
	
	'// 数据库连接字符串（读/写）
	Public Property Get ConnString()
		ConnString = m_ConnString
	End Property
	Public Property Let ConnString(ByVal blParam)
		m_ConnString = blParam
	End Property
	
	'// 得到当前数据库连接接口对象
	Public Property Get Connection()
		IF IsObject(m_Conn) Then Set Connection = m_Conn _
		Else Connect: Set Connection = m_Conn
	End Property

	'// 初始化数据库对象，这个构造器用于程序是直接与数据库建立连接
	Public Sub Connect()
		On Error Resume Next
		
		If IsEmpty(m_Conn) Then
			Set m_Conn = Server.CreateObject("ADODB.Connection")
			m_Conn.ConnectionString = ConnString
			m_Conn.ConnectionTimeout = 15
			m_Conn.Open
		End If

		If Err Then
			m_Conn.Close: Set m_Conn = Nothing
			System.Error.Raise 201
			Response.End()
		End If
		Err.Clear
	End Sub

	'// 关闭并释放数据库连接
	Public Sub DisConnect()
		On Error Resume Next
		If IsObject(m_Conn) Then m_Conn.Close: Set m_Conn = Nothing
		Err.Clear
	End Sub
	
	'// 释放记录集(支持同时释放多个记录集)
	Public Sub DisRecordset(ByVal blObject)
		On Error Resume Next
		If IsArray(blObject) Then
			Dim I: For I = 0 To UBound(blObject)
				If IsObject(blObject(I)) And blObject(I).State = 1 Then blObject(I).Close: Set blObject(I) = Nothing
			Next
		Else
			'// blObject.State=0时，表明数据集为关闭状态
			'// blObject.State=1时，表明数据集为打开状态
			If IsObject(blObject) And blObject.State = 1 Then blObject.Close: Set blObject = Nothing
		End If
		Err.Clear
	End Sub
	Public Sub C(ByVal blObject)
		DisRecordset(blObject)
	End Sub
	
	'/**
	' * @功能说明: 使用参数化查询
	' * @参数说明: - blSource [string]: SQL语句
	' *  		   - blParameters [array]: 参数值。格式：[NAME,TYPE,DIRECTION,SIZE,VALUE],[NAME1,TYPE1,DIRECTION1,SIZE1,VALUE1],[...]
	' * @返回值:   - [recordset] 记录集
	' */
	Public Function Command(ByVal blSQL, ByVal blParameters)
		Dim I, beCommand, beParameter
		Set beCommand = Server.CreateObject("ADODB.Command")
		beCommand.ActiveConnection = Connection
		beCommand.CommandText = blSQL
		beCommand.CommandType = 1
		beCommand.Prepared = True
		
		'// 获取SQL语句字符"?"出现的次数
		Dim blRepeatTimes: blRepeatTimes = System.Text.RepeatTimes("?", blSQL, 0)
		
		Dim blArray, blParamNumber: blParamNumber = 0
		If Not IsArray(blParameters) Then
			'// 当参数为字符串时，将其转换为二维数组
			Set blArray = System.Text.ToArrays(blParameters, "")
			blParamNumber = blArray.Count - 1
		Else blArray = blParameters: blParamNumber = UBound(blArray) End If
		
		If blRepeatTimes = blParamNumber + 1 Then
			For I = 0 To blParamNumber
				Set beParameter = beCommand.CreateParameter
				beParameter.Name      = blArray(I)(0)
				beParameter.Type      = blArray(I)(1)
				beParameter.Direction = blArray(I)(2)
				beParameter.Size      = blArray(I)(3)
				beParameter.Value     = blArray(I)(4)
				beCommand.Parameters.Append beParameter
			Next
			Set Command = beCommand.Execute
			System.Queries = 1
		Else System.Error.Raise 200 End If
	
		Set beCommand = Nothing
	End Function
	
	'/**
	' * @功能说明: 自定义参数，打开记录集
	' * @参数说明: - blSource [string]: SQL语句
	' *  		   - blCursorType [int]: 打开记录集时使用的游标类型
	' *  		   - blLockType [int]: 打开记录集时使用的锁定（并发）类型
	' *  		   - blOptions [int]: 用于指示计算Source参数。如：1为SQL语句，2为表，4为存储过程，8为未知
	' * @返回值:   - [recordset] 记录集
	' */
	Public Function QueryX(ByVal blSource, ByVal blCursorType, ByVal blLockType, ByVal blOptions)
		Dim blRs: Set blRs = Server.CreateObject("ADODB.Recordset")
		blRs.Open blSource, Connection, blCursorType, blLockType, blOptions
		System.Queries = 1
		Set QueryX = blRs
		Set blRs = Nothing
	End Function
	
	'/**
	' * @功能说明: 查询记录
	' * @参数说明: - blSql [string]: SQL查询语句
	' * @返回值:   - [recordset] 记录集
	' */
	Public Function Read(ByVal blSql)
		'SELECT 列名称 FROM 表名称
		Set Read = QueryX(blSql, 1, 1, 1)
	End Function
	
	'/**
	' * @功能说明: 添加记录
	' * @参数说明: - blSql [string]: SQL查询语句
	' * 		  - blArray [array]: 数组，格式：Array(字段名称1:字段值1,字段名称2:字段值2,...)
	' * @返回值:   - [bool] 布尔值
	' */
	Public Function Create(ByVal blSql, ByVal blArray)
		'INSERT INTO 表名称 VALUES (值1, 值2,....)
		'INSERT INTO table_name (列1, 列2,...) VALUES (值1, 值2,....)
		Dim blRs: Set blRs = QueryX(blSql, 1, 2, 1)
		If Not blRs.Bof And Not blRs.Eof Then
			blRs.AddNew
			Dim blCollection, blData: Set blData = System.Array.NewArray(blArray)
			Dim I: For I = 0 To (blData.Size - 1)
				blCollection = System.Text.Separate(blData(I))
				blRs(""&blCollection(0)&"") = blCollection(1)
			Next
			blRs.Update: Create = True
		Else Create = False End If
		blRs.Close: Set blRs = Nothing
	End Function
	
	'/**
	' * @功能说明: 修改记录
	' * @参数说明: - blSql [string]: SQL查询语句
	' * 		   - blArray [array]: 数组，格式：Array(字段名称1:字段值1,字段名称2:字段值2,...)
	' * @返回值:   - [bool] 布尔值
	' */
	Public Function Update(ByVal blSql, ByVal blArray)
		'UPDATE 表名称 SET 列名称 = 新值 WHERE 列名称 = 某值
		Dim blRs: Set blRs = QueryX(blSql, 1, 2, 1)
		If Not blRs.Bof And Not blRs.Eof Then
			Dim blCollection, blData: Set blData = System.Array.NewArray(blArray)
			Dim I: For I = 0 To (blData.Size - 1)
				blCollection = System.Text.Separate(blData(I))
				blRs(""&blCollection(0)&"") = blCollection(1)
			Next
			blRs.Update: Update = True
		Else Update = False End If
		blRs.Close: Set blRs = Nothing
	End Function	
	
	'/**
	' * @功能说明: 删除记录
	' * @参数说明: - blSql [string]: SQL查询语句
	' * @返回值:   - [bool] 布尔值
	' */
	Public Function Delete(ByVal blSql)
		'DELETE FROM 表名称 WHERE 列名称 = 值
		Dim blRs: Set blRs = QueryX(blSql, 1, 2, 1)
		If Not blRs.Bof And Not blRs.Eof Then blRs.Delete: Delete = True _
		Else Delete = False End If
		blRs.Close: Set blRs = Nothing
	End Function
	
	'/**
	' * @功能说明: 根据指定条件生成SQL语句
	' * @参数说明: - blTablePrefix [string]: 数据表名称，格式必须为[表名 字段名1,字段名2,... 取记录的数量]
	' *			  				  [array]: 数据表名称，格式必须为 Array("表名", "字段名1,字段名2,...", "取记录的数量")
	' * 		  - blCondition [string]: 查询条件
	' *			  - blOrderField [string]: 排序方式
	' * @返回值:   - [string] 字符串
	' */
	Public Function ToSQL(ByVal blTablePrefix, ByVal blCondition, ByVal blOrderField)
		Dim blstrSQL, blSqlPrefix		
		Dim blTable, blFields, blTopNumber
		
		With System.Text
			If .IsEmptyAndNull(blTablePrefix) Then ToSQL = "": Exit Function
			Dim blArray: Set blArray = System.Array.New
			blArray.Hash = blTablePrefix 'Format: Array("TABLE", "ID,USER,PASS", "0")
			blTable = blArray(0): blFields = blArray(1)
			blTopNumber = System.Text.ToNumeric(blArray(2))
			Set blArray = Nothing
	
			'// 将参数进行组合成完整的查询语句
			blstrSQL = "Select "
			If blTopNumber > 0 Then blstrSQL = blstrSQL & "Top " & blTopNumber & " "
			blstrSQL = blstrSQL & .IIF(blFields <> "" And blFields <> "*", blFields, "*") & " From " & blTable
			'// 多条件查询，暂时只是将多个条件用AND进行连接
			If Not .IsEmptyAndNull(blCondition) Then
				'If IsArray(blCondition) Then blstrSQL = blstrSQL & " Where " & System.Array.NewArray(blCondition).J(" And ") _
				'Else blstrSQL = blstrSQL & " Where " & blCondition
				blstrSQL = blstrSQL & " Where " & blCondition
			End If
			ToSQL = .IIF(Not .IsEmptyAndNull(blOrderField), (blstrSQL & " Order By " & blOrderField), blstrSQL)
		End With
	End Function
	
	'// 调用存储过程
	Public Function ExecuteSP(ByVal blName, ByVal blParam)
		Dim I, blCommand, blOutParam
		Dim blType: blType = Empty
		
		If GetDataBaseType <> "MSSQL" Then System.Error.Raise 202: Exit Function
		
		If InStr(blName, ":") > 0 Then
			blType = UCase(Trim(Mid(blName, InStr(blName, ":") + 1)))
			blName = Trim(Left(blName, InStr(blName, ":") - 1))
		End If
		
		Set blCommand = Server.CreateObject("ADODB.Command")
		With blCommand
			.ActiveConnection = Connection
			.CommandText = blName
			.CommandType = 4
			.Prepared = True
			.Parameters.Append .CreateParameter("Return", 3, 4)
			blOutParam = "Return"
			
			If Not IsArray(blParam) Then 
				If Not System.Text.IsEmptyAndNull(blParam) Then blParam = System.Text.IIF(InStr(blParam, ",") > 0, sPlit(blParam, ","), Array(blParam))
			End If
			
			If IsArray(blParam) Then
				For I = 0 To UBound(blParam)
					Dim bl_tName, bl_tValue
					If (blType = "1" Or blType = "OUT" Or  blType = "3" Or blType = "ALL") And InStr(blParam(1), "@@") = 1 Then
						.Parameters.Append .CreateParameter(blParam(I), 200, 2, 8000)
						blOutParam = blOutParam & "," & blParam(I)
					Else
						If InStr(blParam(I), "@") = 1 And InStr(blParam(I), ":") > 2 Then
							bl_tName = Left(blParam(I), InStr(blParam(I), ":") - 1)
							blOutParam = blOutParam & "," & bl_tName
							bl_tValue = Mid(blParam(I), InStr(blParam(I), ":") + 1)
							If bl_tValue = "" Then bl_tValue = Null
							.Parameters.Append .CreateParameter(bl_tName, 200, 1, 8000, bl_tValue)
						Else
							.Parameters.Append .CreateParameter("@param" & (I+1), 200, 1, 8000, blParam(I))
							blOutParam = blOutParam & "," & "@param"&(I+1)
						End If
					End If
				Next
			End If
		End With
		
		blOutParam = System.Text.IIF(InStr(blOutParam, ",") > 0, sPlit(blOutParam, ","), Array(blOutParam))
		If blType = "1" Or blType = "OUT" Then
			blCommand.Execute: ExecuteSP = blCommand
		ElseIf blType = "2" Or blType = "RS" Then
			Set ExecuteSP = blCommand.Execute
		ElseIf blType = "3" Or blType = "ALL" Then
			Dim bltOutParam: Set bltOutParam = Server.CreateObject("Scripting.Dictionary")
			Dim bltRs: Set bltRs = blCommand.Execute: bltRs.Close()
			For I = 0 To UBound(blOutParam)
				bltOutParam(Trim(blOutParam(I))) = blCommand(I)
			Next
			bltRs.Open: ExecuteSP = Array(bltRs, bltOutParam)
		Else blCommand.Execute: ExecuteSP = blCommand(0) End If
		Set blCommand = Nothing
	End Function
	
	'// 压缩ACCESS数据库
	Public Sub CompressionAccess()
		Dim JetEngine: Set JetEngine = Server.CreateObject("JRO.JetEngine")
		JetEngine.CompactDatabase ConnString, ConnString &".temp"
		System.IO.FSO.CopyFile Source&".temp", Source
		System.IO.DeleteFile Source&".temp"
		Set JetEngine = Nothing
	End Sub
	
	'// 获取当前使用的数据库类型
	Public Function GetDataBaseType()
		Select Case System.Data.Connection.Provider
			Case "MSDASQL.1", "SQLOLEDB.1", "SQLOLEDB" : GetDataBaseType = "MSSQL"
			Case "MSDAORA.1", "OraOLEDB.Oracle" : GetDataBaseType = "ORACLE"
			Case "Microsoft.Jet.OLEDB.4.0" : GetDataBaseType = "ACCESS"
			Case Else GetDataBaseType = ""
		End Select
	End Function
	
	'// 将记录集转换为JSON格式代码
	'// blParam参数 name[:totalName][:notjs]
	'// name String (字符串) 
	'// 该Json数据在Javascript中的名称 
	'// totalName(可选) String (字符串) 
	'// 如果不省略此参数，则会在生成的Json字符串中添加一个名称为该参数的表示总记录数的项 
	'// notjs(可选) String (字符串) 
	'// 此参数为固定字符串"notjs",如不省略此参数，则输出的Json字符串中不会将中文进行编码 
	Public Function toJSON(ByVal blRs, ByVal blParam)
		'On Error Resume Next
		Dim blField, blTotal		
		Dim blNotJS: blNotJS = False
		Dim blName: blName = System.Text.Separate(blParam)
		
		If Not System.Text.IsEmptyAndNull(blName(1)) Then
			blParam = blName(0): blTotal = blName(1)
			blName = System.Text.Separate(blTotal)
			If Not System.Text.IsEmptyAndNull(blName(1)) Then
				blTotal = blName(0): blNotJS = (LCase(blName(1)) = "notjs")
			End If
		End If
		
		Dim Rs: Set Rs = blRs.Clone
		Dim blCount: blCount = 0
		Dim blJSON: Set blJSON = System.JSON.New(0)
		If blNotJS Then blJSON.StrEncode = False
		If Not System.Text.IsEmptyAndNull(Rs) Then
			blCount = Rs.RecordCount
			If Not System.Text.IsEmptyAndNull(blTotal) Then blJSON(blTotal) = blCount
			blJSON(blParam) = System.JSON.New(1)
			While Not (Rs.EOF Or Rs.BOF)
				blJSON(blParam)(Null) = System.JSON.New(0)
				For Each blField In Rs.Fields
					blJSON(blParam)(Null)(blField.Name) = blField.Value
				Next
				Rs.MoveNext
			Wend
		End If
		toJSON = blJSON.JsString
		Set blJSON = Nothing
		Rs.Close(): Set Rs = Nothing
		If Err.Number <> 0 Then System.Error.Raise 203
		Err.Clear
	End Function
	
End Class

%>

<%
'// +--------------------------------------------------------------------------
'// | Boyle.ACL [系统数据分页操作类]
'// +--------------------------------------------------------------------------
Class Cls_Data_Page
	
	'// 声明私有变量
	Private C
	'// 声明公共变量
	
	'// 初始化资源
	Private Sub Class_Initialize()
		Set C = Dicary(): C.CompareMode = 1
		
		'// 初始化使用的数据库类型
		C("TYPE") = System.Data.GetDataBaseType()
		
		'// 初始化默认分页按钮输出样式
		C("__FIRST") = "&#171;": C("__LAST") = "&#187;"
		C("__PREVIOUS") = "&#8249;": C("__NEXT") = "&#8250;"
		
		'// 初始化分页样式
		C("STYLE") = "PAGER"
		'// 初始化接收当前页的链接标签
		C("LABEL") = "PAGE"
		
		'// 初始化分页所必须的参数
		C("ROWPAGE") = 10: C("PAGESIZE") = 10: C("PAGECOUNT") = 0
		
	End Sub
	
	'// 释放资源
	Private Sub Class_Terminate()
		Set C = Nothing
	End Sub
	
	'// 设置SQL语句
	Public Property Let SQL(ByVal blParam)
		C("SQL") = blParam
	End Property
	Public Property Get SQL()
		SQL = C("SQL")
	End Property
	
	'// 设置参数
	Public Property Let Parameters(ByVal blArray)
		Dim tDic: Set tDic = System.Text.ToHashTable(blArray)
		'Dim tDic: Set tDic = System.Array.NewHash(blArray).Maps
		Dim tKey: For Each tKey In tDic
			C(tKey) = tDic.Item(tKey)
		Next
	End Property
	
	'// 获取参数集合，以JSON方式输出
	Public Property Get Parameters()
		Parameters = System.Text.DictionaryToJSON(C, "PAGEPARAMETERS:NUMBER")
	End Property
	
	'// 获取单一参数
	Public Property Get Parameter(ByVal blItem)
		Parameter = C(blItem)
	End Property
	
	'// 设置分页样式
	Public Property Get Style()
		Style = C("STYLE")
	End Property
	Public Property Let Style(ByVal blParam)
		C("STYLE") = blParam
	End Property
	
	'// 设置地址栏页码标签
	Public Property Get Label()
		Label = C("LABEL")
	End Property
	Public Property Let Label(ByVal blParam)
		C("LABEL") = blParam
	End Property
	
	'// 显示分页码的个数
	Public Property Get rowPage()
		rowPage = C("ROWPAGE")
	End Property
	Public Property Let rowPage(ByVal blSize)
		C("ROWPAGE") = System.Text.ToNumeric(blSize)
	End Property
	'// 每页显示的记录数
	Public Property Get PageSize()
		PageSize = C("PAGESIZE")
	End Property
	Public Property Let PageSize(ByVal blSize)
		C("PAGESIZE") = System.Text.ToNumeric(blSize)
	End Property
	
	'// 总记录数
	Public Property Get RecordCount()
		RecordCount = C("RECORDCOUNT")
	End Property
	Public Property Let RecordCount(ByVal blParam)
		C("RECORDCOUNT") = System.Text.ToNumeric(blParam)
	End Property
	
	'// 总页数
	Public Property Get PageCount()
		PageCount = C("PAGECOUNT")
	End Property
	Public Property Let PageCount(ByVal blParam)
		C("PAGECOUNT") = System.Text.ToNumeric(blParam)
	End Property
	
	'// 当前页码
	Public Property Get CurrentPage()
		Dim tPage: tPage = System.Text.ToNumeric(C("CURRENTPAGE"))
		Dim cPage: cPage = System.Text.ToNumeric(C("PAGECOUNT"))
        If tPage < 1 Then tPage = 1
        If tPage > cPage Then tPage = cPage
		CurrentPage = tPage
	End Property
	Public Property Let CurrentPage(ByVal Param1)
		C("CURRENTPAGE") = System.Text.ToNumeric(Param1)
	End Property
	
	'// 执行分页程序
	Public Function Run()
		Run = Empty
		Dim blRs, blSQL: blSQL = C("SQL")
		Select Case UCase(C("TYPE"))
			Case "1", "MSSQL":
			
			Case "2", "MSSQL-SP":
			Case "3", "MYSQL":
				Run = System.Data.Connection.Execute(blSQL & " LIMIT "& (C("CURRENTPAGE") - 1) * C("PAGESIZE") & "," & C("PAGESIZE")).GetRows()
			Case "4", "ACCESS":
				Set blRs = System.Data.QueryX(blSQL, 1, 1, 1)
				'// 设置总记录数
				C("RECORDCOUNT") = blRs.RecordCount
				'// 设置总页数
				C("PAGECOUNT") = Abs(Int(-(C("RECORDCOUNT") / C("PAGESIZE"))))
				'// 设置当前页
				C("CURRENTPAGE") = CurrentPage
				
				If Not blRs.Bof And Not blRs.Eof Then
					'// ACCESS BUG
					If C("CURRENTPAGE") > 1 And C("CURRENTPAGE") = C("PAGECOUNT") And (C("RECORDCOUNT") Mod C("PAGESIZE") = 1) Then
						blRs.AbsolutePosition = (C("CURRENTPAGE") - 1) * C("PAGESIZE")
					Else blRs.AbsolutePosition = (C("CURRENTPAGE") - 1) * C("PAGESIZE") + 1 End If
					
					Run = blRs.GetRows(C("PAGESIZE"))
				End If
				blRs.Close: Set blRs = Nothing
		End Select
	End Function
	
	'// 输出分页列表
	'// FIRST PREVIOUS 1 2 3 4 5 6 7 8 9 ... 99 100 NEXT LAST PAGER_INFO
	'// PREVIOUS 1 2 3 4 5 6 7 8 9 ... 99 100 NEXT
	'// PREVIOUS 1 2 ... 92 93 94 95 96 97 98 99 100 NEXT
	'// 各种分页样式 http://mis-algoritmos.com/2007/03/16/some-styles-for-your-pagination/
	Public Function Out()
		Dim blHtml: blHtml = Empty
		Dim blUrl: blUrl = GetUrlParam("*", C("LABEL"))
		Dim blListPage, thePage, PrevBound, NextBound
		Dim rowPage: rowPage = System.Text.ToNumeric(C("ROWPAGE"))
		PrevBound = C("CURRENTPAGE") - Int(rowPage / 2)
		NextBound = C("CURRENTPAGE") + Int(rowPage / 2)
		If PrevBound <= 0 Then PrevBound = 1: NextBound = rowPage
		If NextBound > C("PAGECOUNT") Then NextBound = C("PAGECOUNT"): PrevBound = C("PAGECOUNT") - rowPage
		
		If C("PAGECOUNT") = 1 Then
			blHtml = blHtml & "<span class=""current"">1</span>"
		Else
			'// 显示首页和下一页
			If C("CURRENTPAGE") > 1 Then
				Dim blHomeHref: blHomeHref = Replace(blUrl, "*", 1)
				Dim blPreviousHref: blPreviousHref = Replace(blUrl, "*", C("CURRENTPAGE") - 1)
				blHtml = blHtml & "<span><a href="""& blHomeHref &""">"& C("__FIRST") &"</a></span><span><a href="""& blPreviousHref &""">"& C("__PREVIOUS") &"</a></span>"
			Else
				blHtml = blHtml & "<span class=""disabled"">"& C("__FIRST") &"</span><span class=""disabled"">"& C("__PREVIOUS") &"</span>"
			End If
			
			'// 显示页码列表
			For rowPage = PrevBound To NextBound
				If rowPage = C("CURRENTPAGE") Then
					thePage = "<span class=""current"">"& rowPage &"</span>"
				ElseIf rowPage <= C("PAGECOUNT") Then
					thePage = "<span><a href="""& Replace(blUrl, "*", rowPage) &""">"& rowPage &"</a></span>"
				End If
				blListPage = blListPage & thePage
			Next
			blHtml = blHtml & LCase(blListPage)
			
			'// 显示尾页和上一页
			If C("CURRENTPAGE") < C("PAGECOUNT") Then
				Dim blNextHref: blNextHref = Replace(blUrl, "*", C("CURRENTPAGE") + 1)
				Dim blLastHref: blLastHref = Replace(blUrl, "*", C("PAGECOUNT"))
				blHtml = blHtml & "<span><a href="""& blNextHref &""">"& C("__NEXT") &"</a></span><span><a href="""& blLastHref &""">"& C("__LAST") &"</a></span>"
			Else
				blHtml = blHtml & "<span class=""disabled"">"& C("__NEXT") &"</span><span class=""disabled"">"& C("__LAST") &"</span>"
			End If
		End If
		Out = "<div class="""& LCase(C("STYLE")) &""">" & blHtml & "</div>"
	End Function
	
	'// 智能链接组合
	Private Function GetUrlParam(ByVal blPageNumber, ByVal blPageParam)
		Dim blQSItem, blParam: blParam = ""
		For Each blQSItem In Request.QueryString()
			'// 将除指定项除外进行重新拼接
			If UCase(blQSItem) <> UCase(blPageParam) Then
				blParam = blParam & blQSItem & "=" & Server.UrlEncode(Request.QueryString(blQSItem)) & "&"
			End If
		Next
		'// 重组之后，将指定向添加到末尾处
		blParam = "?" & blParam & blPageParam & "=" & blPageNumber
		GetUrlParam = LCase(blParam)
	End Function
End Class

%>
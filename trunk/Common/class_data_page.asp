<%
'// --------------------------------------------------------------------------- //
'// Project Name		: Boyle.ACL												//
'// Program Name		: class_data_page.asp									//
'// Copyright Notice	: COPYRIGHT (C) 2011 BY BOYLE NETWORK.					//
'// Creation Date		: 2011/08/02											//
'// Version				: 3.1.0.0802											//
'//																				//
'// Date       By			 Description										//
'// ---------- ------------- -------------------------------------------------- //
'// 2011/08/02 Boyle	 	 系统数据分页操作类									//
'// --------------------------------------------------------------------------- //

Class Cls_Data_Page
	
	'// 声明私有变量
	Private C
	'// 声明公共变量
	
	'// 初始化资源
	Private Sub Class_Initialize()
		Set C = Server.CreateObject("Scripting.Dictionary")
		C.CompareMode = 1
		
		'// 初始化使用的数据库类型
		C("TYPE") = GetDataBaseType()
		
		'// 初始化默认分页按钮输出样式
		C("__FIRST") = "&#171;": C("__LAST") = "&#187;"
		C("__PREVIOUS") = "&#8249;": C("__NEXT") = "&#8250;"
		
		'// 初始化分页样式
		C("STYLE") = "PAGER"
		'// 初始化接收当前页的链接标签
		C("URLPARAM") = "PAGE"
		
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
		Dim tKey: For Each tKey In tDic
			C(tKey) = tDic.Item(tKey)
		Next
	End Property
	
	'// 获取参数集合，以JSON方式输出
	Public Property Get Parameters()
		Parameters = System.Text.DictionaryToJSON(C, "PAGEPARAMETERS", 0)
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
	Public Property Get UrlParam()
		UrlParam = C("URLPARAM")
	End Property
	Public Property Let UrlParam(ByVal blParam)
		C("URLPARAM") = blParam
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
        If tPage < 1 Then tPage = 1
        If tPage > C("PAGECOUNT") Then tPage = System.Text.ToNumeric(C("PAGECOUNT"))		
		CurrentPage = tPage
	End Property
	Public Property Let CurrentPage(ByVal Param1)
		Dim tPage: tPage = System.Text.ToNumeric(Param1)
		If tPage < 1 Then tPage = 1
		If tPage > C("PAGECOUNT") Then tPage = C("PAGECOUNT")
		C("CURRENTPAGE") = System.Text.ToNumeric(tPage)
	End Property
	
	'// 执行分页程序
	Public Function Run()
		Run = Empty
		Dim blRs, blSQL
		Select Case UCase(C("TYPE"))
			Case "1", "MSSQL":
			
			Case "2", "MSSQL-SP":
			Case "3", "MYSQL":
				Run = System.Data.Connection.Execute(C("SQL") & " LIMIT "& (C("CURRENTPAGE") - 1) * C("PAGESIZE") & "," & C("PAGESIZE")).GetRows()
			Case "4", "ACCESS":
				Set blRs = System.Data.QueryX(C("SQL"), 1, 1, 1)
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
		Dim blUrl: blUrl = GetUrlParam("*", C("URLPARAM"))
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
			If UCase(blQSItem) <> blPageParam Then
				blParam = blParam & blQSItem & "=" & Server.UrlEncode(Request.QueryString(blQSItem)) & "&"
			End If
		Next
		'// 重组之后，将指定向添加到末尾处
		blParam = "?" & blParam & blPageParam & "=" & blPageNumber
		GetUrlParam = LCase(blParam)
	End Function
	
	'// 获取当前使用的数据库类型
	Private Function GetDataBaseType()
		Select Case System.Data.Connection.Provider
			Case "MSDASQL.1", "SQLOLEDB.1", "SQLOLEDB" : GetDataBaseType = "MSSQL"
			Case "MSDAORA.1", "OraOLEDB.Oracle" : GetDataBaseType = "ORACLE"
			Case "Microsoft.Jet.OLEDB.4.0" : GetDataBaseType = "ACCESS"
			Case Else GetDataBaseType = ""
		End Select
	End Function
End Class

%>
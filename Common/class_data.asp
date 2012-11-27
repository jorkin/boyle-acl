<!--#include file="./class_data_page.asp"-->
<%
'// --------------------------------------------------------------------------- //
'// Project Name		: Boyle.ACL												//
'// Author				: Boyle(boyle7[at]qq.com)								//
'// Copyright Notice	: COPYRIGHT (C) 2011-2012 BY BOYLE.						//
'// Create Date			: 2011/08/02											//
'// Version				: 4.0.121028											//
'//																				//
'// Date       By			 Description										//
'// ---------- ------------- -------------------------------------------------- //
'// 2011/08/02 Boyle	 	 系统数据库操作类									//
'// --------------------------------------------------------------------------- //

Class Cls_Data

	'// 声明私有对象
	Private PrPager
	Private m_Conn, m_Source, m_ConnString
	
	'// 声明公共对象

	'// 初始化资源
	Private Sub Class_Initialize()
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
	
	'/**
	' * 数据库地址（读/写）
	' * @params 数据库地址
	' * @return 数据库地址（字符串）
	' */	
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
			If Not m_Conn Is Nothing Then
				Response.Write "ADO Version: " & m_Conn.Version & "<br />"
				If m_Conn.State = &H00000001 Then m_Conn.Close
			End If
			Set m_Conn = Nothing
			Response.Write(Err.Source & "<br />" & Err.Description)
			Err.Clear: Response.End()
		End If
	End Sub

	'// 关闭连接，如果是直接与数据库连接的情况则什么也不做
	'// 如果是从连接池中取得的连接那么释放传来的连接
	Public Sub DisConnect()
		On Error Resume Next
		If IsObject(m_Conn) Then m_Conn.Close: Set m_Conn = Nothing
		Err.Clear
	End Sub
	
	'// 释放记录集(支持同时释放多个记录集)
	Public Sub DisRecordset(ByVal blObject)
		Dim I: If IsArray(blObject) Then
			For I = 0 To UBound(blObject)
				If IsObject(blObject(I)) And blObject(I).State = 1 Then blObject(I).Close: Set blObject(I) = Nothing
			Next
		Else
			'// blObject.State=0时，表明数据集为关闭状态
			'// blObject.State=1时，表明数据集为打开状态
			If IsObject(blObject) And blObject.State = 1 Then blObject.Close: Set blObject = Nothing
		End If
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
		Else Set Command = Nothing: System.Error.Throw "提示:错误的参数个数!" End If
	
		' System.Log.X = Array("SQL:" & blSQL)
		' System.Log.X = Array("1:" & beCommand.Parameters(0) & "-Type:" & beCommand.Parameters(0).Type)
		' System.Log.X = Array("2:" & beCommand.Parameters(1) & "-Type:" & beCommand.Parameters(1).Type)
		' System.Log.X = Array("3:" & beCommand.Parameters(2) & "-Type:" & beCommand.Parameters(2).Type)
		' System.Log.X = Array("4:" & beCommand.Parameters(3) & "-Type:" & beCommand.Parameters(3).Type)
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
		Set QueryX = blRs
		Set blRs = Nothing
	End Function
	
	'/**
	' * @功能说明: 查询记录
	' * @参数说明: - blSql [string]: SQL查询语句
	' * @返回值:   - [recordset] 记录集
	' */
	Public Function Query(ByVal blSql)
		Set Query = QueryX(blSql, 1, 1, 1)
	End Function
	
	'/**
	' * @功能说明: 添加记录
	' * @参数说明: - blSql [string]: SQL查询语句
	' * 		   - blArray [array]: 必须为二维数据集，格式：[字段名称1, 字段值1][字段名称2, 字段值2]...
	' * @返回值:   - [bool] 布尔值
	' */
	Public Function Insert(ByVal blSql, ByVal blArray)
		Dim blRs: Set blRs = QueryX(blSql, 1, 2, 1)
		If Not blRs.Bof And Not blRs.Eof Then
			blRs.AddNew
			Dim I: For I = 0 To UBound(blArray)
				blRs(""& blArray(I)(0) &"") = blArray(I)(1)
			Next
			blRs.Update: Insert = True
		Else Insert = False End If
		blRs.Close: Set blRs = Nothing
	End Function
	
	'/**
	' * @功能说明: 修改记录
	' * @参数说明: - blSql [string]: SQL查询语句
	' * 		   - blArray [array]: 必须为二维数据集，格式：[字段名称1, 字段值1][字段名称2, 字段值2]...
	' * @返回值:   - [bool] 布尔值
	' */
	Public Function Modify(ByVal blSql, ByVal blArray)
		Dim blRs: Set blRs = QueryX(blSql, 1, 2, 1)
		If Not blRs.Bof And Not blRs.Eof Then
			Dim I: For I = 0 To UBound(blArray)
				blRs(""& blArray(I)(0) &"") = blArray(I)(1)
			Next
			blRs.Update: Modify = True
		Else Modify = False End If
		blRs.Close: Set blRs = Nothing
	End Function	
	
	'/**
	' * @功能说明: 删除记录
	' * @参数说明: - blSql [string]: SQL查询语句
	' * @返回值:   - [bool] 布尔值
	' */
	Public Function Delete(ByVal blSql)
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
			blTable = blArray(0)
			blFields = .IIF(Not .IsEmptyAndNull(blArray(1)), .ToString(.ToArray(blArray(1), ","), "],["), blArray(1))
			blTopNumber = System.Text.ToNumeric(blArray(2))
			Set blArray = Nothing
	
			'// 将参数进行组合成完整的查询语句
			blstrSQL = "Select "
			If blTopNumber > 0 Then blstrSQL = blstrSQL & "Top " & blTopNumber & " "
			blstrSQL = blstrSQL & .IIF(blFields <> "" And blFields <> "*", "[" & blFields & "]", "* ") & " From [" & blTable & "]"
			'// 多条件查询，暂时只是将多个条件用AND进行连接
			If IsArray(blCondition) Then blstrSQL = blstrSQL & " Where " & .ToString(blCondition, " And ")_
			Else blstrSQL = .IIF(blCondition <> "", (blstrSQL & " Where " & blCondition), blstrSQL)
			ToSQL = .IIF(blOrderField <> "", (blstrSQL & " Order By " & blOrderField), blstrSQL)
		End With
	End Function
	
	'// 调用存储过程
	Public Function ExecuteSP(ByVal blName, ByVal blParam)
		Dim I, blCommand, blOutParam
		Dim blType: blType = Empty
		
		If IsEmptyAndNull(DataBaseType) And Not UCase(DataBaseType) = "MSSQL" And DataBaseType <> 1 Then
			System.WE ("仅支持从MS SQL Server数据库中调用存储过程！"): Exit Function
		End If
		
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
				If Not IsEmptyAndNull(blParam) Then blParam = System.Text.IIF(InStr(blParam, ",") > 0, sPlit(blParam, ","), Array(blParam))
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
		If Err.Number <> 0 Then System.Error.Raise "生成Json格式代码出错！"
	End Function
	
End Class

%>
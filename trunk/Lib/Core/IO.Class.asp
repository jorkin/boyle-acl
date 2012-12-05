<!--#include file="./IO.CAS.class.asp"-->
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
'// 2012/12/29 Boyle		 系统文件操作类										//
'// --------------------------------------------------------------------------- //

Class Cls_IO
	
	'// 定义私有命名对象
	Private PrFSO, PrCAS
	
	'// 定义FSO名称，以及操作文件的编码
	Private PrName, PrCharset
	
	'// 初始化类
	Private Sub Class_Initialize()
		'// 初始化FSO对象名称
		PrName = "Scripting.FileSystemObject"
		
		'// 初始化编码格式，继承自主类
		PrCharset = System.Charset
		
		'// 初始化系统默认的错误信息
		System.Error.E(52) = "写入文件错误！"
		System.Error.E(53) = "创建文件夹错误！"
		System.Error.E(54) = "读取文件列表失败！"
		System.Error.E(55) = "设置属性失败，文件不存在！"
		System.Error.E(56) = "设置属性失败！"
		System.Error.E(57) = "获取属性失败，文件不存在！"
		System.Error.E(58) = "复制失败，源文件不存在！"
		System.Error.E(59) = "移动失败，源文件不存在！"
		System.Error.E(60) = "删除失败，文件不存在！"
		System.Error.E(61) = "重命名失败，源文件不存在！"
		System.Error.E(62) = "重命名失败，已存在同名文件！"
		System.Error.E(63) = "文件或文件夹操作错误！"
	End Sub
	
	'// 释放类
	Private Sub Class_Terminate()
		If IsObject(PrFSO) Then Set PrFSO = Nothing End If
		If IsObject(PrCAS) Then Set PrCAS = Nothing End If
	End Sub
	
	'// 声明文件操作对象模块单元
	Public Property Get FSO()
		If Not IsObject(PrFSO) Then Set PrFSO = Server.CreateObject(PrName) End If
		Set FSO = PrFSO
	End Property
	Public Property Get CAS()
		If Not IsObject(PrCAS) Then Set PrCAS = New Cls_IO_CAS End If
		Set CAS = PrCAS
	End Property
	
	'// 设置服务器FSO组件的名称
	Public Property Let Name(ByVal blParam)
		PrName = blParam
	End Property
	
	'// 设置操作文件的编码
	Public Property Let Charset(ByVal blParam)
		PrCharset = blParam
	End Property
	
	'/**
	' * @功能说明: 动态包含文件
	' * @参数说明: - blFilePath [string]: 目标文件路径
	' */
	Public Sub Import(ByVal blFilePath)
		ExecuteGlobal ReadInclude(blFilePath, 0)
	End Sub
	
	'// 获取文件内容
	Private Function ReadInclude(ByVal blFilePath, ByVal blHtml)
		Dim blContentStartPosition, blCodeStartPosition
		Dim blContent, blTempContent, blCode, blTempCode, blHtmlCode
		blContent = ReadIncludes(blFilePath)
		blCode = "": blContentStartPosition = 1: blCodeStartPosition = InStr(blContent, "<"&"%") + 2
		blHtmlCode = System.Text.IIF(blHtml = 1, "blACLHtml = blACLHtml & ","Response.Write ")
		While blCodeStartPosition > blContentStartPosition + 1
			blTempContent = Mid(blContent, blContentStartPosition, blCodeStartPosition - blContentStartPosition - 2)
			blContentStartPosition = InStr(blCodeStartPosition, blContent, "%"&">") + 2
			If Not System.Text.IsEmptyAndNull(blTempContent) Then
				blTempContent = Replace(blTempContent, """", """""")
				blTempContent = Replace(blTempContent, vbCrLf&vbCrLf, vbCrLf)
				blTempContent = Replace(blTempContent, vbCrLf, """&vbCrLf&""")
				blCode = blCode & blHtmlCode & """" & blTempContent & """" & vbCrLf
			End If
			blTempContent = Mid(blContent, blCodeStartPosition, blContentStartPosition - blCodeStartPosition - 2)
			blTempCode = System.Text.ReplaceX(blTempContent, "^\s*=\s*", blHtmlCode) & vbCrLf
			If blHtml = 1 Then
				blTempCode = System.Text.ReplaceXMultiline(blTempCode, "^(\s*)Response\.Write", "$1" & blHtmlCode) & vbCrLf
				blTempCode = System.Text.ReplaceXMultiline(blTempCode, "^(\s*)System\.(WB|W|WE|WR)", "$1" & blHtmlCode) & vbCrLf
			End If
			blCode = blCode & Replace(blTempCode, vbCrLf&vbCrLf, vbCrLf)
			blCodeStartPosition = InStr(blContentStartPosition, blContent, "<"&"%") + 2
		Wend
		blTempContent = Mid(blContent,blContentStartPosition)
		If Not System.Text.IsEmptyAndNull(blTempContent) Then
			blTempContent = Replace(blTempContent, """", """""")
			blTempContent = Replace(blTempContent, vbCrLf&vbCrLf, vbCrLf)
			blTempContent = Replace(blTempContent, vbCrlf,"""&vbCrLf&""")
			blCode = blCode & blHtmlCode & """" & blTempContent & """" & vbCrLf
		End If
		If blHtml = 1 Then blCode = "blACLHtml = """" " & vbCrLf & blCode
		ReadInclude = Replace(blCode, vbCrLf&vbCrLf, vbCrLf)
	End Function
	
	'// 递归获取包含文件的内容
	Private Function ReadIncludes(ByVal blFilePath)
		Dim blContent: blContent = Me.Read(blFilePath)
		If Not System.Text.IsEmptyAndNull(blContent) Then
			blContent = System.Text.ReplaceX(blContent, "<"&"% *?@.*?%"&">", "")
			blContent = System.Text.ReplaceX(blContent, "(<"&"%[^>]+?)(option +?explicit)([^>]*?%"&">)", "$1'$2$3")
			Dim blRule: blRule = "<!-- *?#include +?(file|virtual) *?= *?""??([^"":?*\f\n\r\t\v]+?)""?? *?-->"
			'// 判断文件中是否有包含其他文件
			If System.Text.Test(blContent, blRule) Then
				Dim blIncludeFile, blIncludeFileContent
				Dim blMatches: Set blMatches = System.Text.MatchX(blContent, blRule)
				Dim blMatch: For Each blMatch In blMatches
					If LCase(blMatch.SubMatches(0)) = "virtual" Then blIncludeFile = blMatch.SubMatches(1) _
					Else blIncludeFile = Mid(blFilePath, 1, InstrRev(blFilePath, System.Text.IIF(Instr(blFilePath, ":") > 0, "\", "/"))) & blMatch.SubMatches(1)
					'// 递归获取包含文件的内容
					blIncludeFileContent = ReadIncludes(blIncludeFile)
					blContent = Replace(blContent, blMatch, blIncludeFileContent)
				Next
				Set blMatches = Nothing
			End If
		End If
		ReadIncludes = blContent
	End Function
		
	'/**
	' * @功能说明: 读取指定文件对象
	' * @参数说明: - blFile [string]: 对象路径
	' * @返回值:   - [string] 字符串
	' */
	Public Function Read(ByVal blFile)
		If ExistsFile(blFile) Then
			blFile = FormatFilePath(blFile)
			Dim blStream: Set blStream = Server.CreateObject("ADODB.Stream")
			With blStream
				.Type = 2: .Mode = 3
				.Charset = PrCharset: .Open
				.LoadFromFile blFile: Read = .ReadText
				.Close
			End With
			Set blStream = Nothing
		Else Read = "" End If
	End Function

	'/**
	' * @功能说明: 覆盖当前打开的文本内容，文件及文件夹不存在则创建
	' * @参数说明: - blFile [string]: 对象路径
	' * 		   - blContent [string]: 保存内容
	' * @返回值:   - [bool]: 布尔值
	' */
	Public Function Save(ByVal blFile, ByVal blContent)
		If Not System.Text.IsEmptyAndNull(blFile) Then
			blFile = FormatFilePath(blFile)
			Dim blFolder: blFolder = Directory(blFile, "\")

			'// 如果文件夹不存在，则创建新文件夹
			If Not ExistsFolder(blFolder) Then CreateFolder(blFolder)
			'// 只有在文件夹存在时，对文件进行保存
			If ExistsFolder(blFolder) Then
				On Error Resume Next
				Dim blStream: Set blStream = Server.CreateObject("ADODB.Stream")
				With blStream
					.Open
					.Charset = PrCharset
					.Position = blStream.Size
					.WriteText = blContent
					.SaveToFile blFile, 2
					.Close
				End With
				If Err Then
					System.Error.Message = "（"& blFile &"）"
					System.Error.Raise 52
				End If
				Err.Clear
				Save = True: Set blStream = Nothing
			Else Save = False End If
		Else Save = False End If
	End Function
	
	'/**
	' * @功能说明: 删除文件(同时支持绝对和相对两种路径模式)
	' * @参数说明: - blFile [string]: 删除文件对象
	' * @返回值:   - [bool] 布尔值
	' */
	Public Function Delete(ByVal blFile)		
		If Not ExistsFile(blFile) Then Delete = False _
		Else FSO.DeleteFile FormatFilePath(blFile): Delete = True
	End Function
	
	'/**    
	' * @功能说明： 遍历目录下的所有目录和文件（不包括子目录）
	' * @参数说明： - [string] blPath : 初始路径    
	' * 			- [bool] blShowFile : 是否遍历文件
	' * @返回值：   - [array] : 二维数组
	' */	
	Public Function Dir(ByVal blPath, ByVal blShowFile)
		If ExistsFolder(blPath) Then
			Dim blArray(), I: I = 0
			Dim blFolder, blSubFolder, blItem
			Set blFolder = FSO.GetFolder(FormatFilePath(blPath))
			Set blSubFolder = blFolder.SubFolders
			ReDim Preserve blArray(4, blSubFolder.Count - 1)
			For Each blItem In blSubFolder
				blArray(0, I) = blItem.Name & "/"
				blArray(1, I) = blItem.Size
				blArray(2, I) = blItem.DateLastModified
				blArray(3, I) = blItem.Attributes
				blArray(4, I) = blItem.Type
				I = I + 1
			Next
			'// 判断是否显示文件
			If System.Text.ToBoolean(blShowFile) Then
				Set blSubFolder = blFolder.Files
				ReDim Preserve blArray(4, blSubFolder.Count + I - 1)
				For Each blItem In blSubFolder
					blArray(0, I) = blItem.Name
					blArray(1, I) = blItem.Size
					blArray(2, I) = blItem.DateLastModified
					blArray(3, I) = blItem.Attributes
					blArray(4, I) = blItem.Type
					I = I + 1
				Next
			End If
			Set blSubFolder = Nothing
			Set blFolder = Nothing
			Dir = blArray
		Else ReDim blArray2(-1, -1): Dir = blArray2 End If
	End Function
	
	'/**    
	' * @功能说明： 遍历目录下的所有目录和文件（包括子目录）
	' * @参数说明： - [string] sPath : 初始路径    
	' *  			- [bool] bAll : 是否遍历子目录
	' * 			- [bool] bFile : 是否遍历文件
	' * @返回值：   - [array] : 数组
	' */    
	Public Function Traversal(ByVal sPath, ByVal bAll, ByVal bFile)
		If ExistsFolder(sPath) Then
			Dim oKey, pKey, nKey, mKey
			Dim mItem, mName, nItem, nPath
			Dim oDic, oArray
			Set oDic = Server.CreateObject("Scripting.Dictionary")
			For Each nItem In FSO.GetFolder(FormatFilePath(sPath)).SubFolders
				nPath = sPath & nItem.Name & "/"
				oKey = System.Security.MD5(nPath, 16)
				If Not oDic.Exists(oKey) Then oDic.Add oKey, nPath
				If System.Text.ToBoolean(bFile) Then
					For Each mItem In nItem.Files
						mName = nPath & mItem.Name
						nKey = System.Security.MD5(mName, 16)
						If Not oDic.Exists(nKey) Then oDic.Add nKey, mName
					Next
				End If
				
				If System.Text.ToBoolean(bAll) Then
					If System.Text.ToBoolean(bFile) Then oArray = Traversal(nPath, True, True) _
					Else oArray = Traversal(nPath, True, False)
					
					Dim I: For I = 0 To UBound(oArray)
						pKey = System.Security.MD5(oArray(I), 16)
						If Not oDic.Exists(pKey) Then oDic.Add pKey, oArray(I)
						If System.Text.ToBoolean(bFile) Then
							For Each mItem In nItem.Files
								mName = nPath & mItem.Name
								mKey = System.Security.MD5(mName, 16)
								If Not oDic.Exists(mKey) Then oDic.Add mKey, mName
							Next
						End If
					Next
				End If
			Next
			Traversal = oDic.Items
			Set oDic = Nothing
		Else ReDim blArray(-1): Traversal = blArray End If
	End Function

	'/**
	' * @功能说明: 新建文件夹（同时支持绝对和相对两种路径模式）
	' * @参数说明: - blFolder [string]: 新建文件夹对象
	' * @返回值:   - [bool] 布尔值
	' */
	Public Function CreateFolder(ByVal blFolder)
		If Not System.Text.IsEmptyAndNull(blFolder) Then
			Dim I, ArrayPaths
			Dim blPaths: blPaths = Split(FormatFilePath(blFolder), "\")
			For I = 0 To UBound(blPaths)
				If I = 0 Then ArrayPaths = blPaths(I) Else ArrayPaths = ArrayPaths & "\" & blPaths(I)
				If I > 0 Then
					'// 当前文件夹下，如果有文件与文件夹同名时，将无法创建文件夹。
					If ExistsFile(ArrayPaths) Then
						System.Error.Message = "("& ArrayPaths &")"
						System.Error.Raise 53
						CreateFolder = False: Exit Function
					Else
						If Not ExistsFolder(ArrayPaths) Then FSO.CreateFolder ArrayPaths End If
					End If
				End If
			Next
			CreateFolder = True
		Else CreateFolder = False End If
	End Function

	'/**
	' * @功能说明: 删除文件夹(同时支持绝对和相对两种路径模式)
	' * @参数说明: - blFolder [string]: 删除文件夹对象
	' * @返回值:   - [bool] 布尔值
	' */
	Public Function DeleteFolder(ByVal blFolder)		
		If Not System.Text.IsEmptyAndNull(blFolder) Then
			If ExistsFolder(blFolder) Then FSO.DeleteFolder FormatFilePath(blFolder): DeleteFolder = True _
			Else DeleteFolder = False
		Else DeleteFolder = False End If
	End Function

	'/**
	' * @功能说明: 检查文件夹是否存在
	' * @参数说明: - blFolder [string]: 对象路径
	' * @返回值:   - [bool] 布尔值
	' */
	Public Function ExistsFolder(ByVal blFolder)
		If System.Text.IsEmptyAndNull(blFolder) Then ExistsFolder = False _
		Else ExistsFolder = FSO.FolderExists(FormatFilePath(blFolder))
	End Function

	'/**
	' * @功能说明: 检查文件是否存在
	' * @参数说明: - blFile [string]: 检测对象的路径
	' * @返回值:   - [bool] 布尔值
	' */
	Public Function ExistsFile(ByVal blFile)
		If System.Text.IsEmptyAndNull(blFile) Then ExistsFile = False _
		Else ExistsFile = FSO.FileExists(FormatFilePath(blFile))
	End Function

	'/**
	' * @功能说明: 将目标文件（相对路径）转换为绝对路径
	' * @参数说明: - blFile [string]: 目标文件路径
	' * @返回值:   - [string] 字符串
	' * @函数说明: 函数先转换所出现的错误字符。基于：Errorchar变量
	' *            判断当前输入参数是绝对路径或相对路径
	' *            如果是绝对路径，替换所有的相对路径所使用的字符
	' */
	Public Function FormatFilePath(ByVal blFile)
		Dim I, blParam3: blParam3 = Empty
		Dim blParam2: blParam2 = Trim(blFile)
		Dim blIllegal: blIllegal = "',"",*,?,&,|,<,>,;"
		Dim blArrayIllegal: blArrayIllegal = Split(blIllegal, ",")		
		For I = 0 To UBound(blArrayIllegal)
			If InStr(blParam2, blArrayIllegal(I)) > 0 Then
				blParam2 = Replace(blParam2, blArrayIllegal(I), "")
			End If
		Next
		
		blParam2 = System.Text.ReplaceX(blParam2, "(\/|\\)+", "/")
		
		'// 判断目标路径是否为绝对路径
		If Mid(blParam2, 2, 1) <> ":" Then blParam2 = Server.MapPath(blParam2) _
		Else blParam2 = Replace(blParam2, "/", "\")
		
		FormatFilePath = System.Text.IIF(Right(blParam2, 1) = "\", Left(blParam2, Len(blParam2) - 1), blParam2)
	End Function

	'// 取文件夹绝对路径
	Private Function absPath(ByVal p)
		If System.Text.IsEmptyAndNull(p) Then absPath = "" : Exit Function
		If Mid(p, 2, 1) <> ":" Then
			If isWildcards(p) Then
				p = Replace(p, "*", "[.$.[a.c.l.s.t.a.r].#.]")
				p = Replace(p, "?", "[.$.[a.c.l.q.u.e.s].#.]")
				p = Server.MapPath(p)
				p = Replace(p, "[.$.[a.c.l.q.u.e.s].#.]", "?")
				p = Replace(p, "[.$.[a.c.l.s.t.a.r].#.]", "*")
			Else
				p = Server.MapPath(p)
			End If
		End If
		If Right(p, 1) = "\" Then p = Left(p, Len(p) - 1)
		absPath = p
	End Function
	
	'// 路径是否包含通配符
	Private Function isWildcards(ByVal path)
		isWildcards = False
		If InStr(path, "*") > 0 Or InStr(path, "?") > 0 Then isWildcards = True
	End Function
	
	'/**
	' * @功能说明: 获取指定文件的目录
	' * @参数说明: - blFile [string]: 目标文件路径
	' * 		   - blParam2 [string]: 查询关键字，默认为"/"
	' * @返回值:   - [string]: 字符串
	' */
	Public Function Directory(ByVal blFile, ByVal blParam2)
		Dim blSplits: blSplits = System.Text.IIF(System.Text.IsEmptyAndNull(blParam2), "/", blParam2)
		Directory = System.Text.IIF(InStrRev(blFile, blSplits) < 1, "/", Mid(blFile, 1, InStrRev(blFile, blSplits)))
	End Function

	'/**
	' * @功能说明: 获取文件的后缀名
	' * @参数说明: - blFile [string]: 目标文件
	' * @返回值:   - [string] 字符串
	' */
	Public Function FileExts(ByVal blFile)
		FileExts = "Unknow"
		FileExts = LCase(Split(blFile, ".")(UBound(Split(blFile, "."))))
	End Function	

	'// 设置文件或文件夹属性
	Public Function [Attributes](ByVal path, ByVal attrType)
		On Error Resume Next
		Dim p,a,i,n,f,at : p = Me.FormatFilePath(path) : n = 0 : [Attributes] = True
		
		If Not ExistsFile(P) Or Not ExistsFolder(P) Then
			[Attributes] = False
			System.Error.Message = "(" & path & ")"
			System.Error.Raise 55
			Exit Function
		End If
		
		If ExistsFile(p) Then
			Set f = FSO.GetFile(p)
		ElseIf ExistsFolder(p) Then
			Set f = FSO.GetFolder(p)
		End If
		at = f.Attributes : a = UCase(attrType)
		If Instr(a,"+")>0 Or Instr(a,"-")>0 Then
			a = System.Text.IIF(Instr(a," ")>0, Split(a," "), Split(a,","))
			For i = 0 To Ubound(a)
				Select Case a(i)
					Case "+R" at = System.Text.IIF(at And 1,at,at+1)
					Case "-R" at = System.Text.IIF(at And 1,at-1,at)
					Case "+H" at = System.Text.IIF(at And 2,at,at+2)
					Case "-H" at = System.Text.IIF(at And 2,at-2,at)
					Case "+S" at = System.Text.IIF(at And 4,at,at+4)
					Case "-S" at = System.Text.IIF(at And 4,at-4,at)
					Case "+A" at = System.Text.IIF(at And 32,at,at+32)
					Case "-A" at = System.Text.IIF(at And 32,at-32,at)
				End Select
			Next
			f.Attributes = at
		Else
			For i = 1 To Len(a)
				Select Case Mid(a,i,1)
					Case "R" n = n + 1
					Case "H" n = n + 2
					Case "S" n = n + 4
				End Select
			Next
			f.Attributes = System.Text.IIF(at And 32,n+32,n)
		End If
		Set f = Nothing
		If Err.Number <> 0 Then
			[Attributes] = False
			System.Error.Message = "(" & path & ")"
			System.Error.Raise 56
		End If
		Err.Clear()
	End Function
	
	'// 获取文件或文件夹信息
	Public Function GetAttributes(ByVal path, ByVal attrType)
		Dim f,s,p : p = Me.FormatFilePath(path)
		If ExistsFile(p) Then
			Set f = FSO.GetFile(p)
		ElseIf ExistsFolder(p) Then
			Set f = FSO.GetFolder(p)
		Else
			GetAttributes = ""
			System.Error.Message = "(" & path & ")"
			System.Error.Raise 57
			Exit Function
		End If
		Select Case LCase(attrType)
			Case "0","name" : s = f.Name
			Case "1","date", "datemodified" : s = f.DateLastModified
			Case "2","datecreated" : s = f.DateCreated
			Case "3","dateaccessed" : s = f.DateLastAccessed
			Case "4","size" : s = FormatSize(f.Size, s_sizeformat)
			Case "5","attr" : s = Attr2Str(f.Attributes)
			Case "6","type" : s = f.Type
			Case Else s = ""
		End Select
		Set f = Nothing
		GetAttributes = s
	End Function	
	
	'// 格式化文件大小
	Public Function FormatSize(Byval fileSize, ByVal level)
		Dim s : s = Int(fileSize) : level = UCase(level)
		FormatSize = System.Text.IIF(s/(1073741824)>0.01,FormatNumber(s/(1073741824),2,-1,0,-1),"0.01") & " GB"
		If s = 0 Then FormatSize = "0 GB"
		If level = "G" Or (level="AUTO" And s>1073741824) Then Exit Function
		FormatSize = System.Text.IIF(s/(1048576)>0.1,FormatNumber(s/(1048576),1,-1,0,-1),"0.1") & " MB"
		If s = 0 Then FormatSize = "0 MB"
		If level = "M" Or (level="AUTO" And s>1048576) Then Exit Function
		FormatSize = System.Text.IIF((s/1024)>1,Int(s/1024),1) & " KB"
		If s = 0 Then FormatSize = "0 KB"
		If Level = "K" Or (level="AUTO" And s>1024) Then Exit Function
		If level = "B" or level = "AUTO" Then
			FormatSize = s & " bytes"
		Else
			FormatSize = s
		End If
	End Function
	
	'// 格式化文件属性
	Private Function Attr2Str(ByVal attrib)
		Dim a,s : a = Int(attrib)
		If a>=2048 Then a = a - 2048
		If a>=1024 Then a = a - 1024
		If a>=32 Then : s = "A" : a = a- 32 : End If
		If a>=16 Then a = a- 16
		If a>=8 Then a = a - 8
		If a>=4 Then : s = "S" & s : a = a- 4 : End If
		If a>=2 Then : s = "H" & s : a = a- 2 : End If
		If a>=1 Then : s = "R" & s : a = a- 1 : End If
		Attr2Str = s
	End Function
	
End Class
%>
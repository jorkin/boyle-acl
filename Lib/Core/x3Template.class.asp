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
'// 2011/08/02 Boyle		 系统模板操作类										//
'// --------------------------------------------------------------------------- //

'// --------------------------------------------------------------------------- //
'// 作者：Taihom(taihom@163.com)(Taihom.Template.class v3.0)					//
'// 网址：http://www.cnblogs.com/taihom/										//
'// --------------------------------------------------------------------------- //

Class Cls_Template
	'// 声明私有变量
	Private strCharset, strSITEROOT, strTemplate_dir, strTemplate_path, strCache_dir
	Private strTemplate, strCachePage, strCachePageName
	Private intCacheflag, intAbsPath, intCachePageTimeout, intCreateCachePage, intCacheTplTime
	Private DIC_BLOCK_ATTR, DIC_BLOCK_LOOP_VAL, DIC_BLOCK_IF, DIC_BLOCK_LOOP, DIC_BLOCK_LOOP_LIST, DIC_BLOCK_LOOP_ATTR
	Private PREFIX, RS_FIELD, OPS
	
	'// 常用对象
	Private FSO, XMLDOM, EX
	
	'// 类初始化
	Private Sub Class_Initialize()		
		'// 全局默认变量
		strCharset          = System.Charset
		strSITEROOT         = ""'Server.MapPath("./")
		intCacheflag        = 0 	 '是否需要缓存 0=不缓存,1=文件缓存,2=内存
		strCache_dir        = ""	 '后面需要带/
		strTemplate_dir     = ""	 '后面需要带/
		intAbsPath          = 1 	 '输出结果是否使用绝对路径 (0不用,1用)
		intCachePageTimeout = 0 	 '整个页面的缓存超时
		intCacheTplTime     = 121028 '模板缓存时间
		
		'// 以下信息请不要更改除非你对本类的思路跟实现比较了解否则请不要更改
		PREFIX = "GLOBAL_BOYLE_TEMPLATE"
		'// 默认支持的运算符
		OPS = "==|<=|&lt;=|>=|&gt;=|>|&gt;|<|&lt;|!=|<>|&lt;&gt;|%"
		
		'// 常用对象
		Set DIC_BLOCK_ATTR      = Dicary()	'加入的名值对象
		Set DIC_BLOCK_IF        = Dicary()	'if块的对象
		Set DIC_BLOCK_LOOP      = Dicary()	'loop块的对象
		Set DIC_BLOCK_LOOP_VAL  = Dicary()	'循环体内容的值
		Set DIC_BLOCK_LOOP_ATTR = Dicary()	'块属性输出
		Set DIC_BLOCK_LOOP_LIST = Dicary()	'块输出对象
		Set RS_FIELD            = Dicary()	'应用到loop中的字段名

		Set XMLDOM = Server.CreateObject("MicroSoft.XMLDom")
		XMLDOM.async = False
		
		Set FSO = System.IO.FSO
		Set EX = New RegExp
		EX.Global = True
		EX.IgnoreCase = True
	End Sub
	
	'// 类退出
	Private Sub Class_Terminate()
		'// 释放字典
		Set DIC_BLOCK_ATTR      = Nothing
		Set DIC_BLOCK_IF        = Nothing
		Set DIC_BLOCK_LOOP      = Nothing
		Set DIC_BLOCK_LOOP_VAL  = Nothing
		Set DIC_BLOCK_LOOP_ATTR = Nothing
		Set DIC_BLOCK_LOOP_LIST = Nothing
		Set RS_FIELD            = Nothing
		'// 释放对象
		Set XMLDOM = Nothing  
		Set EX     = Nothing
		Set FSO    = Nothing
	End Sub

	'// 字典对象
	Private Function Dicary()
		Set Dicary = Server.CreateObject("Scripting.Dictionary")
	End Function
	
	'// 循环嵌套include(从代码解析,来源路径)
	Private Function LoadInclude(ByVal html, ByVal fromPath)
		Dim incpath
		Dim Match, Matches
		'// include优先
		EX.Pattern = "{include\(['""]*(.+?)['""]*\)}"
		Set Matches = EX.Execute(html)
		For Each Match in Matches
			'incpath = Replace(strSITEROOT & strTemplate_dir & Match.SubMatches(0), "/", "\")
			incpath = strSITEROOT & strTemplate_dir & Match.SubMatches(0)
			If LCase(fromPath) <> LCase(incpath) Then
				html = Replace(html, Match.Value, LoadInclude(LoadFile(incpath, strCharset), incpath))
			Else html = Replace(html, Match.Value,"") End If
		Next
		Set Matches = Nothing
		LoadInclude = html
	End Function
	
	'// 读取模板文件
	Private Function LoadFile(ByVal strFilePath, ByVal strCharset)
		If System.IO.ExistsFile(strFilePath) Then LoadFile = System.IO.Read(strFilePath) _
		Else LoadFile = "模板"&strFilePath&"加载失败"
	End Function
	
	'// 加载模板文件
	Private Function LoadTemplate()
		Dim html, tplpath: tplpath = strSITEROOT & strTemplate_dir & strTemplate_path
		'// 加载模板，如果为空，则直接读取内存，否则读取模板文件
		If Len(strTemplate_path) = 0 Then html = strTemplate _
		Else html = LoadFile(tplpath, strCharset)
		'// include优先:LoadInclude(代码,来源路径)
		html = LoadInclude(html, tplpath)
		
		'// 把模板内容存储到临时变量，是否替换模板内容中的相对路径为绝对路径
		If Len(html) Then strTemplate = System.Text.IIF(CBool(intAbsPath), AbsPath(html), html)		
	End Function
	
	'// 保存文件
	Private Function SaveToFile(ByVal strContent, ByVal strFilePath, ByVal strCharset)
		SaveToFile = System.IO.Save(strFilePath, strContent)
	End Function
	
	'// FSO自动生成文件夹路径
	Private Function AutoCreateFolder(ByVal strPath)
		AutoCreateFolder = System.IO.CreateFolder(strPath)
	End Function
	
	'// 输出结果输出模板的绝对路径
	Private Function AbsPath(ByVal strCode)
		Dim html: html = strCode
		Dim Matches, Match
		EX.Pattern = "(href|src)=(['""|])(?!(\/|\{|\(|\.\/|http:\/\/|https:\/\/|javascript:|#))(.+?)(['""|])"
		Set Matches = EX.Execute(html)
		For Each Match in Matches
			html = Replace( html, Match.Value, Replace(Match.value, Match.SubMatches(3), RelPath(Match.SubMatches(3))) )
		Next
		AbsPath = html
		'AbsPath = EX.Replace(html, "$1=$2" & Replace(strTemplate_dir,"\","/") & "$3$4$5") 
	End Function
		
	'// 替换相对路径，根据模板路径把../逐层替换到对应的目录
	Private Function RelPath(ByVal strPath)
		Dim src, spath
		EX.Pattern = "^(\.\.\/)+"
		Dim Matches: Set Matches = EX.Execute(strPath)
		If Matches.Count = 0 Then RelPath = Replace(strTemplate_dir, "\", "/") & strPath: Exit Function
		'// 模板的全路径
		spath = Split(Replace(strTemplate_dir & strTemplate_path, "\", "/"), "/")
		EX.Pattern = "(\.\.\/)" '//看有多少个../
		Dim I: For I = 0 To Ubound(spath) - 1 - EX.Execute(Matches(0).Value).Count
			src = src & spath(I) & "/"
		Next
		'// 把../替换成正确的目录
		RelPath = Replace(strPath, Matches(0).Value, src)
	End Function

	'// 标签属性的支持，如果需要做自己的属性支持，在这里扩展
	Private Function TagProperty(ByVal strTag, ByVal tagVal, ByVal strVal)
		Dim return: return = strVal
		tagVal = Trim(tagVal)
		Select Case (LCase(Trim(strTag)))
			'// 截取长度的支持
			Case "len", "length":
				Dim tagVals: tagVals = System.Text.Separate(tagVal)
				If Int(tagVals(0)) > 0 Then return = System.Text.Cut(return, tagVals(0)&":"&tagVals(1))
			'// 返回值{@var return="len"}
			Case "return":
				Select Case (LCase(Trim(tagVal)))
					Case "len", "length":'// 返回变量的字符串长度
						return = System.Text.Length(return)
					Case "urlencode":'// 返回URLEncode
						return = Server.URLEncode(return)
					Case "clearhtml":'// 清除html格式
						return = System.Text.RemoveHtml(return)
					Case "clearspace":'// 清除所有空格空行
						return = System.Text.RemoveSpace(return)
					Case "clearformat":'// 清除所有格式
						return = System.Text.RemoveSpace(System.Text.RemoveHtml(return))
				End Select
			'// 日期格式化
			Case "dateformat":
				return = DateFormat(return, tagVal)		
		End Select
		TagProperty = return
	End Function

	'// 日期格式化
	Public Property Get DateFormat(ByVal strDate, ByVal strFormat)
		Dim return: return = strFormat
		If Not IsDate(strDate) Then DateFormat = strDate : Exit Property		
		'下面开始进行日期的转换
		Dim ccDate: ccDate = FormatDateTime(strDate, vbGeneralDate )
		'显示方式
		Select Case(strFormat)
			Case "0" : 'vbGeneralDate 0 显示日期和/或时间。如果有日期部分，则将该部分显示为短日期格式。如果有时间部分，则将该部分显示为长时间格式。如果都存在，则显示所有部分。
				return = FormatDateTime(strDate, 0 )
			Case "1" : 'vbLongDate 1 使用计算机区域设置中指定的长日期格式显示日期
				return = FormatDateTime(strDate, 1 )
			Case "2" : 'vbShortDate 2 使用计算机区域设置中指定的短日期格式显示日期 
				return = FormatDateTime(strDate, 2 )
			Case "3" : 'vbLongTime 3 使用计算机区域设置中指定的时间格式显示时间
				return = FormatDateTime(strDate, 3 )
			Case "4" : 'vbShortTime 4 使用 24 小时格式 (hh:mm) 显示时间
				return = FormatDateTime(strDate, 4 )
			Case Else
				With EX
					'年月日是小写，时分秒是大写
					.IgnoreCase = False
					.Global = True
					.Pattern = "yyyy" : return = .Replace(return, Year(ccDate))
					.Pattern = "yy"   : return = .Replace(return, Right(Year(ccDate), 2))
					.Pattern = "mm"   : return = .Replace(return, Right("0" & Month(ccDate), 2))
					.Pattern = "m"    : return = .Replace(return, Month(ccDate))
					.Pattern = "dd"   : return = .Replace(return, Right("0" & Day(ccDate), 2))
					.Pattern = "d"    : return = .Replace(return, Day(ccDate))
					.Pattern = "HH"   : return = .Replace(return, Right("0" & Hour(ccDate), 2))
					.Pattern = "H"    : return = .Replace(return, Hour(ccDate))
					.Pattern = "MM"   : return = .Replace(return, Right("0" & Minute(ccDate), 2))
					.Pattern = "M"    : return = .Replace(return, Minute(ccDate))
					.Pattern = "SS"   : return = .Replace(return, Right("0" & Second(ccDate), 2))
					.Pattern = "S"    : return = .Replace(return, Second(ccDate))
					.Pattern = "w|W"  : return = .Replace(return, Right(weekdayname(weekday(ccDate)), 1))
					.IgnoreCase = True
				End With
		End Select
		
		DateFormat = return
	End Property
	
	'////////////////////////////////////////////////模板输出部分'////////////////////////////////////////////////

	'// 判断输出，返回真或者假
	Private Function ifOperator(ByVal strOpera, ByVal strVar, ByVal strValue)
		On Error Resume Next
		Dim return : return = False		
		strVar = Trim(strVar): strValue = Trim(strValue)		
		'// 运算符支持
		Select Case(LCase(strOpera))
			Case "==":
				return = (CStr(strVar) = CStr(strValue))
			Case "<=", "&lt;=":
				return = (CDbl(strVar) <= CDbl(strValue))
			Case ">=", "&gt;=":
				return = (CDbl(strVar) >= CDbl(strValue))
			Case ">", "&gt;":
				return = (CDbl(strVar) > CDbl(strValue))
			Case "<", "&lt;":
				return = (CDbl(strVar) < CDbl(strValue))
			Case "!=", "<>", "&lt;&gt;":
				return = (CStr(strVar) <> CStr(strValue))
			Case "%", "mod":'// 模运算支持,感谢优能科技提供代码
				strValue = Split(strValue, "=")
				return = (CLng(strVar) mod CLng(strValue(0)) = CLng(strValue(1)))
		End Select		
		ifOperator = System.Text.IIF((Err.Number = 0), return, False): Err.Clear
	End Function
	
	'// 单值替换,替换Assign_Tag标签
	Private Function Assign_tag(ByVal strCode, ByVal Pattern)
		Dim html : html = strCode
		Dim Match, Matches		
		'// 单标签处理
		EX.Pattern = Pattern
		Set Matches = EX.Execute(html)		
		'// 单标签支持
		'// 0=标签名, 1=标签属性
		For Each Match in Matches
			html = Replace(html, Match.value, Return_Tag_Val(DIC_BLOCK_ATTR(LCase(Match.SubMatches(0))), Match.SubMatches(1)))
		Next		
		'// 在html中添加额外的meta
		'EX.Pattern = "<head>" : Set Matches = EX.Execute(html)
		'For Each Match in Matches: html = Replace(html, Match.Value, Match.value&vbnewline&strMeta): Next
		Set Matches = Nothing
		Assign_tag = html
	End Function
	
	'// ASP函数调用，可以调用asp的内置函数或者自定义函数
	Private Function Assign_asp(ByVal strCode)
		On Error Resume Next
		Dim html: html = strCode
		Dim Match, Matches
		'// 单标签处理
		EX.Pattern = "\{asp\s+(.+?)\s?\}"
		Set Matches = EX.Execute(html)		
		'// 单标签支持
		'// 0=标签名, 1=标签属性
		For Each Match In Matches: html = Replace(html, Match.Value, Eval(Trim(Match.SubMatches(0)))): Next
		Set Matches = Nothing: Err.Clear
		Assign_asp = html
	End Function
	
	'// 逻辑判断块的输出
	Private Function Assign_if(ByVal strCode)
		Dim html: html = strCode
		Dim return: return = False
		Dim strVar, strOperator, strValue
		Dim OutputTag, arrBlock, block_index
		Dim Match, Matches

		EX.Pattern = "{if name=""(.*?)""}"
		Set Matches = EX.Execute(html)
		For Each Match in Matches
			block_index = Trim(Match.SubMatches(0))'if块的索引
			'// 获取到索引后,读取字典数据
			'// 0=索引,1=条件(),2=分支
			arrBlock = DIC_BLOCK_IF(block_index)			
			'if的条件
			strVar      = Assign_tag(arrBlock(1)(0), "(@\w+)(.*)?")	'// 外部输入的值
			strOperator = Trim(arrBlock(1)(1))						'// 运算符
			strValue    = Assign_tag(arrBlock(1)(2), "(@\w+)(.*)?")	'// 支持表达式两边是变量			
			'// 表达式运算
			return = ifOperator(strOperator, strVar, strValue)			
			'// 分支的选择
			'// 如果有多个分支
			If Ubound(arrBlock(2)) > 0 Then OutputTag = System.Text.IIF(return, arrBlock(2)(0), arrBlock(2)(1)) _
			Else OutputTag = System.Text.IIF(return, arrBlock(2)(0), "")			
			'// 输出标签
			html = Replace(html, Match.Value, OutputTag)		
		Next
		Set Matches = Nothing
		Assign_if = html
	End Function
	
	'// 提取出loop里面的empty标签，并且解析出loop标签
	'// 返回数组0=loop主体,1=empty
	Private Function Assign_loop_bodytag(ByVal strCode)
		Dim aryBody(1), Match, Matches
		EX.Pattern = "<empty>([\s\S]*?)</empty>"
		Set Matches = EX.Execute(strCode)
		'// 如果有<empty>标签
		If Matches.Count > 0 Then
			For Each Match In Matches
				'// 提取主体
				aryBody(0) = EX.Replace(strCode, "")
				aryBody(1) = Match.SubMatches(0)
			Next
		Else aryBody(0) = strCode: aryBody(1) = "" End If
		Set Matches = Nothing
		Assign_loop_bodytag = aryBody
	End Function
	
	'// 循环块的输出
	Private Function Assign_loop(ByVal strCode)
		Dim html: html = strCode
		Dim BlockValue, OutputTag
		Dim block_index, arrBlock, strBlock, subBlock
		Dim Match, Matches
		
		'// 循环体输出<loop>
		EX.Pattern = "{loop name=""(.*?)""}"
		Set Matches = EX.Execute(html)
		For Each Match in Matches
			block_index = Trim(Match.SubMatches(0))'loop块的索引
			'// 获取到索引后,读取字典数据
			'// 0索引,1='数组对象(0=名,1=值),2循环体
			arrBlock = DIC_BLOCK_LOOP(block_index)			
			'// 块的值
			BlockValue = DIC_BLOCK_LOOP_VAL(block_index)			
			'// 块的循环体
			'// 检测循环体中的<empty>标签,并且提取出来
			strBlock = Assign_loop_bodytag(arrBlock(2))			
			OutputTag = ""
			'// 如果输入的是数组
			If IsArray(BlockValue) Then
				Set subBlock = GetLoopBlock(arrBlock(2)) '// 获得循环块				
				Dim rs: rs = BlockValue
				'// 循环输出记录集
				'// 由于是二维数组，没有记录到字段名，所以数据输出只能用下标
				Dim I: For I = 0 To Ubound(rs, 2)
					'// 标签输出(块索引,循环内容,记录集,序号)
					OutputTag = OutputTag & Assign_loop_Tag(block_index, block_index, strBlock(0), rs, I)
					'// 嵌套循环支持
					IF TypeName(subBlock)="IMatchCollection2" Then OutputTag = Assign_subloop(block_index,OutputTag)
					'// 序号输出
					OutputTag = Replace(OutputTag, "(i)", I + 1)
					'// 支持loop里面的if
					OutputTag = Return_If_Tag_Val(OutputTag)
				Next
			Else
				OutputTag = strBlock(1)	'// 支持自定义记录为空的模板
				OutputTag = Return_If_Tag_Val(OutputTag)	'// 支持loop里面的if
			End If			
			html = Replace(html, Match.Value, OutputTag)
		Next
		Set Matches = Nothing
		Assign_loop = html
	End Function
	
	'// 解析子循环部分
	Private Function Assign_subloop(ByVal block_index,ByVal strCode)
		Dim html : html = strCode
		Dim Match, Matches
		Dim attr, arrAttr
		Dim subLoop, subAttr, subBody, strBlock, blockid
		Dim rs, I, substr
		Set Matches = GetLoopBlock(strCode)'获得对象
		For Each Match in Matches'0=标签名,1=参数部分(如果没有就为空),2=循环体
			subLoop = Replace(Match.SubMatches(0), ":", "")'标签名
			subAttr = Match.SubMatches(1)'属性部分
			subBody = Match.SubMatches(2)'循环体
			arrAttr = Analysis_block_attr(subAttr)'子循环的标签名跟属性组
			
			'//索引规则是 上一级的 loopname.本级.loopname loop>loop
			blockid = block_index&"."&subLoop
			
			'// 块的循环体
			'// 检测循环体中的<empty>标签,并且提取出来
			'strBlock = Assign_loop_bodytag(subBody)
			
			'如果有属性组
			If IsArray(arrAttr) Then'获得子循环的值并且替换输出
				rs = GetSubLoopData(blockid, subAttr)'获得数据
				substr = ""
				If TypeName(rs) = "Variant()" Then'二维数据
					'循环输出记录集
					'由于是二维数组，没有记录到字段名，所以数据输出只能用下标
					For I = 0 To Ubound(rs, 2)
						'// 标签输出(块内容名称,块索引,循环内容,记录集,序号)
						substr = substr & Assign_loop_Tag(subLoop, blockid, subBody, rs, I)
						substr = Replace(substr, "(ii)", I + 1)'序号输出
						substr = Return_If_Tag_Val(substr)'支持loop里面的if
					Next
					html = Replace(html, Match.Value, substr)
				'如果数据为空
				Else html = Replace(html, Match.Value, "") End If
				'清空数据集
				If DIC_BLOCK_LOOP_VAL.Exists(blockid) Then DIC_BLOCK_LOOP_VAL.Remove(blockid)
			Else html = Match.value End If
		Next
		Set Matches = Nothing
		Assign_subloop = html
	End Function
	
	'// 获得子循环中的块的值(属性参数部分)
	Private Function GetSubLoopData(ByVal blockid,ByVal subAttr)
		On Error Resume Next
		Dim data: data = Empty
		Dim attr, rs, I, strTag: strTag = blockid
		attr = GetAttribute(subAttr, "data") '// 支持 data 属性
		If Not IsEmpty(attr) Then
			Set rs = Eval(attr) '// 执行
			If TypeName(rs)="Recordset" Then
				'字段设置-----------在这里建立循环中的对应关系
				For I = 0 To rs.Fields.Count - 1 
					RS_FIELD(strTag&"."&LCase(rs.Fields(I).name)) = I '// 建立字段索引
				Next
				'// 赋值到循环块列表中
				If Not rs.Eof Then
					data = rs.GetRows '// 返回getrows数据
					If DIC_BLOCK_LOOP_VAL.Exists(strTag) Then DIC_BLOCK_LOOP_VAL.Remove(strTag)
					DIC_BLOCK_LOOP_VAL(strTag) = data '//赋值到循环块列表中
				End If
				'// 程序后面还可以引用rs记录集，所以这里不用关闭，不过在外面的时候记得要关闭
				'rs.Close()
			End If
		End If
		GetSubLoopData = data:Err.Clear
	End Function
	
	'// 在属性中获得单标签的参数(从什么字符串中,检测什么参数名)
	Private Function GetAttribute(ByVal attrCode, ByVal attrName)
		Dim Matches, attr: attr = Empty
		'// 获得某个参数参数
		EX.Pattern = "\s?("&attrName&")\s*=\s*['""|]([\s\S.]*?)['""|]\s+"
		Set Matches = EX.Execute(" " & attrCode & " ")		
		If Matches.Count Then attr = Matches(0).SubMatches(1)
		Set Matches = Nothing
		GetAttribute = attr
	End Function
	
	'// 循环块内部的处理,替换循环体的标签：块名,块索引,循环体,记录集,行号
	Private Function Assign_loop_Tag(ByVal strTag, ByVal strTagIndex, ByVal strBody, ByVal arrRs, ByVal intRowsNum)
		Dim OutputTag, html: html = strBody
		Dim id, rs: rs = arrRs
		Dim Match, Matches

		'EX.Pattern = "\(@([\d+|\w+]+)(.*?)\)"
		EX.Pattern = "\((@|"&strTag&"\.)([\d|\w]+)(.*?)\)"
		EX.Global = True
		EX.IgnoreCase = True
		Set Matches = EX.Execute(strBody) '执行搜索
		For Each Match in Matches
			OutputTag = ""
			'1=字段名或者下标,2=输出属性
			'Ubound(rs,1)字段大小
			id = Match.SubMatches(1)'索引出列名或者列id
			'循环体的标签名是字段名还是下标，最终是以下标为准，如果传递的是字段名，也要把索引ID返回
			id = System.Text.IIF(IsNumeric(id), id, RS_FIELD(strTagIndex&"."&LCase(id)))
			
			'单标签的值
			'如果下标超过字段宽度，那么这个值为空
			If Len(id) > 0 Then
				id = Int(id)
				OutputTag = System.Text.IIF( (id > Ubound(rs, 1)), "", Return_Tag_Val(rs(id, intRowsNum), Match.SubMatches(2)) )
			End If			
			html = Replace(html, Match.Value, OutputTag)
		Next
		Set Matches = Nothing
		'// 一条完整的循环内容返回
		Assign_loop_Tag = html
	End Function
	
	'// 支持loop里的if
	Private Function Return_If_Tag_Val(ByVal strBody)
		Dim OutputTag, tmpStr, tmpArr, html: html = strBody
		Dim return, strVar, strOperator, strValue
		Dim Match, Matches
		
		EX.Pattern = "{if\s*(.+?)\s*("&OPS&")\s*['""|]*(.+?)['""|]*\s*}([\s\S.]*?){\/if}"
		Set Matches = EX.Execute(strBody) '执行搜索
		'0=条件,1=表达式,2值,3=分支
		For Each Match in Matches
			strVar = Assign_tag(Trim(Match.SubMatches(0)),"(@\w+)(.*)?")
			strOperator = Match.SubMatches(1)
			strValue = Assign_tag(Trim(Match.SubMatches(2)),"(@\w+)(.*)?")
			tmpArr = Split(Match.SubMatches(3), "{else}")'分支的选择
			
			If Ubound(tmpArr) Then tmpStr = tmpArr(1)
			'如果表达式成立,返回分支1，否则返回分支2
			OutputTag = System.Text.IIF(ifOperator(strOperator, strVar, strValue), tmpArr(0), tmpStr)
			
			html = Replace(html, Match.Value, OutputTag)
		Next
		Set Matches = Nothing
		Return_If_Tag_Val = html
	End Function
	
	'// 单标签替换输出,支持带属性输出
	Private Function Return_Tag_Val(ByVal strVal, ByVal strAttr)
		'传递值和属性参数过来，根据属性 返回不同的值回去
		If IsNull(strVal) Then Return_Tag_Val="": Exit Function
		Dim arrAttr, tagVal
		Dim return: return = strVal
		arrAttr = Analysis_block_attr(Trim(strAttr))'获取属性列表
		'如果有参数
		IF IsArray(arrAttr) Then
			Dim I: For I = 0 To Ubound(arrAttr)'逐个检查属性
				'arrAttr(I,0)'标签
				'arrAttr(I,1)'输入的值
				return = TagProperty(Trim(arrAttr(I, 0)), arrAttr(I,1), return)'标签属性的支持
			Next
		End If
		Return_Tag_Val = return
	End Function
	
	'// 变量替换
	Private Function AssignTpl()
		Dim html: html = strTemplate		
		html = Assign_if(html)					'先逻辑
		html = Assign_loop(html)				'循环块
		html = Assign_tag(html,"{(@\w+)(.*?)}")	'单标签
		html = Assign_asp(html)					'ASP函数调用支持
		strTemplate = html
	End Function
	
	'////////////////////////////////////////////////模板输出部分'////////////////////////////////////////////////
	
	'////////////////////////////////////////////////模板解析部分'////////////////////////////////////////////////
	
	'// 解析loop 块的属性,返回属性的数组
	Private Function Analysis_block_attr(ByVal strCode)
		Dim arr: arr = False
		Dim Match, Matches
		Dim I: I = 0
		
		EX.Pattern = "\s?(\w.+?)\s*=\s*['""|]([\s\S.]*?)['""|]\s+"
		Set Matches = EX.Execute(" " & strCode & " ")
		
		If Matches.Count Then
			ReDim arr(Matches.Count-1, 1)
			'0=属性,1=属性值
			For Each Match In Matches
				arr(I, 0) = Match.SubMatches(0)
				arr(I, 1) = Match.SubMatches(1)
				I = I + 1
			Next
		End If
		Set Matches = Nothing
		Analysis_block_attr = arr
	End Function
	
	'// 解析 if 块
	Private Function Analysis_block_if(ByVal strCode)
		Dim html: html = strCode
		Dim Match, Matches
		Dim block_id, block_name
		Dim arrAttr(2), arrBranch '分支结构
		Dim I, J: J = 0
		Dim arrBlock_if(2)
		
		EX.Pattern = "{if\s*(.+?)\s*("&OPS&")\s*['""|]*(.+?)['""|]*\s*}([\s\S.]*?){\/if}"
		Set Matches = EX.Execute(html)
		'0=条件,1=表达式,2值,3=分支
		For Each Match in Matches
			J = J + 1 '序号作为块的标识
			'0=条件,1=表达式,2值,3=分支
			arrAttr(0) = Match.SubMatches(0)
			arrAttr(1) = Match.SubMatches(1)
			arrAttr(2) = Match.SubMatches(2)
			arrBranch = Match.SubMatches(3)
			
			block_id = "block_" & J
			block_name = "{if name=""@name""}" '块的名字
			block_name = Replace(block_name, "@name", block_id)
			
			'缓存块id
			arrBlock_if(0) = block_id'索引
			arrBlock_if(1) = arrAttr'条件
			arrBlock_if(2) = Split(arrBranch , "{else}")'分支结构
			
			'把块变量保存到块的数据字典
			DIC_BLOCK_IF(block_id)=arrBlock_if			
			'把if块替换成单标签,并且有特定的标记
			html = Replace(html, Match.Value, block_name)
		Next
		Set Matches = Nothing
		Analysis_block_if = html
	End Function
	
	'// 解析循环模块
	Private Function Analysis_block_loop(ByVal strCode)
		Dim html: html = strCode
		Dim Match, Matches
		Dim block_id, block_name
		Dim arrAttr, arrBranch '分支结构
		Dim subLoop, subAttr
		Dim I, J: J = 0
		Dim K: K = 0
		Dim arrBlock_loop(2)
		
		Set Matches = GetLoopBlock(strCode) '执行搜索。
		'0=标签名,1=参数部分(如果没有就为空),2=循环体
		If TypeName(Matches) <> "IMatchCollection2" Then Analysis_block_loop = strCode: Exit Function End If
		
		For Each Match in Matches
			J = J + 1 '序号作为块的标识
			block_name = Replace(Trim(Match.SubMatches(0)), ":", "")
			arrAttr = Analysis_block_attr(Match.SubMatches(1))
			If IsArray(arrAttr) Then
				block_id = System.Text.IIF(Len(block_name), block_name, "block_" & J)
				block_name = "{loop name=""@name""}" '块的名字
				block_name = Replace(block_name , "@name" , block_id)
				
				'缓存块id
				arrBlock_loop(0) = block_id
				arrBlock_loop(1) = arrAttr'数组对象，0=名,1=值
				arrBlock_loop(2) = Match.SubMatches(2)'循环体
				
				'把块变量保存到块的数据字典
				DIC_BLOCK_LOOP(block_id) = arrBlock_loop
				DIC_BLOCK_LOOP_LIST(block_id) = arrAttr'输出块内容，输出值是:块名，属性对象（数组结构）
				
				'设置输出属性
				For K = 0 To Ubound(arrAttr)
					DIC_BLOCK_LOOP_ATTR(block_id & "." & arrAttr(K, 0)) = arrAttr(K, 1)
				Next		
			Else block_name = Match.Value End If
			
			'把if块替换成单标签,并且有特定的标记
			html = Replace(html, Match.value, block_name)
		Next
		Set Matches = Nothing
		Analysis_block_loop = html
	End Function
	
	'// 获得循环块
	Private Function GetLoopBlock(ByVal strCode)		
		EX.Pattern = "<loop(:[[a-zA-Z0-9]+)?\s*?([\s\S.]*?)>([\s\S.]*?)<\/loop\1>"
		Dim Matches: Set Matches = EX.Execute(strCode)
		If Matches.Count Then Set GetLoopBlock = Matches _
		Else Set GetLoopBlock = Nothing
		Set Matches = Nothing
	End Function
	
	'// 解析单标签,主要是获取属性值,这里不做返回值
	Private Function Analysis_block_tag(ByVal strCode, ByVal Pattern)
		Dim html: html = strCode
		Dim Match, Matches, K, TagName, arrAttr
		
		'单标签处理
		EX.Pattern = Pattern
		Set Matches = EX.Execute(html) '执行搜索
		
		'单标签支持
		'0=标签名,1=标签属性
		For Each Match in Matches
			TagName = Trim(LCase(Match.SubMatches(0)))'0=标签名
			arrAttr = Analysis_block_attr(Match.SubMatches(1))'1=标签属性
			'设置输出属性
			If IsArray(arrAttr) Then
				For K = 0 To Ubound(arrAttr)
					DIC_BLOCK_LOOP_ATTR(TagName & "." & arrAttr(K, 0)) = arrAttr(K, 1)
				Next
			End If
		Next
		Set Matches = Nothing
		'仅仅是解析标签不用返回值
	End Function
	
	'// 模板html解析
	Private Sub Analysis_Html()
		Dim html: html = strTemplate		
		'解析@单标签(属性支持,不返回值)
		call Analysis_block_tag(html, "{(@\w+)(.*?)}")

		html = Analysis_block_loop(html)	'先解析loop块
		html = Analysis_block_if(html)		'解析if块
		strTemplate = html 					'解析后的html 放回临时变量中	
	End Sub
	
	'// 解析Xml模板
	Private Sub Analysis_Xml()
		Dim ifNode, loopNode
		Dim I, J
		Dim arrBlock(2)
		
		'读取配置文件
		XMLDOM.load(strCacheDir & "_.xml")
		
		Dim block_id, block_attr, block_tpl
		Dim array_1(2), array_2, array_3, array_4
		Set ifNode = XMLDOM.SelectSingleNode("//template/if")'if有多少个节点
		For I = 0 To ifNode.ChildNodes.Length - 1		
			block_id = ifNode.ChildNodes(I).GetAttribute("name") '节点name属性
			Set block_attr = ifNode.ChildNodes(I).getElementsByTagName("attr").item(0)
			array_1(0) = block_attr.GetAttribute("var")
			array_1(1) = block_attr.GetAttribute("operator")
			array_1(2) = block_attr.GetAttribute("value")			
			'分支
			Set block_tpl = ifNode.ChildNodes(I).getElementsByTagName("tpl")
			ReDim array_2(block_tpl.Length - 1)
			For J = 0 To block_tpl.Length - 1: array_2(J) = block_tpl.item(J).text: Next			
			arrBlock(0) = block_id
			arrBlock(1) = array_1
			arrBlock(2) = array_2			
			'把块变量保存到块的数据字典
			DIC_BLOCK_IF(block_id) = arrBlock
		Next
		
		'解析loop
		'Redim arrBlock(2)
		Dim element
		Set loopNode = XMLDOM.SelectSingleNode("//template/loop")'loop有多少个节点
		For I = 0 To loopNode.ChildNodes.Length - 1		
			block_id = loopNode.ChildNodes(I).GetAttribute("name") '节点name属性
			Set block_attr = loopNode.ChildNodes(I).getElementsByTagName("attr").item(0).attributes
			ReDim array_3(block_attr.Length - 1, 1)
			'遍历这个集合
			J = 0
			For Each element In block_attr
				array_3(J,0) = element.nodename'属性名
				array_3(J,1) = element.nodevalue'属性值
				DIC_BLOCK_LOOP_ATTR(block_id & "." & array_3(J,0)) = array_3(J,1)'设置输出属性
				J = J + 1
			Next

			'分支
			Set block_tpl = loopNode.ChildNodes(I).getElementsByTagName("tpl").item(0)
			array_4 = block_tpl.text
			
			arrBlock(0) = block_id
			arrBlock(1) = array_3
			arrBlock(2) = array_4
			
			'把块变量保存到块的数据字典
			DIC_BLOCK_LOOP(block_id) = arrBlock
			DIC_BLOCK_LOOP_LIST(block_id) = array_3'输出块内容，输出值是:块名，属性对象（数组结构）
		Next
	End Sub
		
	'// 解析
	Private Sub Analysis()
		If (intCacheflag = 1) Then
			call Analysis_block_tag(strTemplate, "{(@\w+)(.*?)}")
			call Analysis_Xml()
		Else call Analysis_Html() End If
	End Sub
	'////////////////////////////////////////////////模板解析部分'////////////////////////////////////////////////
	
	'////////////////////////////////////////////////模板缓存部分'////////////////////////////////////////////////
	
	'// 把模板的属性配置缓存到xml文件，变量传递的是缓存位置
	Private Function Analysis_SaveToXml(ByVal strPath)
		Dim path: path = strPath		
		Dim keys, items, I
		Dim tplNode, blockNode, block_attr, block_code		
		'创建一个节点对象
		Set tplNode = XMLDOM.CreateElement("template")
		'保存这个节点对象
		XMLDOM.appendChild tplNode
		XMLDOM.SelectSingleNode("//template").appendChild(XMLDOM.CreateElement("if"))'添加一个 if 节点
		XMLDOM.SelectSingleNode("//template").appendChild(XMLDOM.CreateElement("loop"))'添加一个 loop 节点
		
		'// 解析if块到xml
		For Each keys in DIC_BLOCK_IF
			items = DIC_BLOCK_IF(keys)			
			'添加一个ifblock节点
			Set blockNode = XMLDOM.CreateElement("block")
			blockNode.SetAttribute "name", keys
			XMLDOM.SelectSingleNode("//template/if").appendChild(blockNode)			
			'创建一个attr节点
			XMLDOM.SelectSingleNode("//template/if/block[@name='"&keys&"']").appendChild(XMLDOM.CreateElement("attr"))			
			'获取if 的attr属性	
			block_attr = items(1)
			'0=变量，1=运算符,2=对比值
			For I = 0 To Ubound(block_attr)
				'设置attr的属性
				XMLDOM.SelectSingleNode("//template/if/block[@name='"&keys&"']/attr").SetAttribute "var", block_attr(0)'变量
				XMLDOM.SelectSingleNode("//template/if/block[@name='"&keys&"']/attr").SetAttribute "operator", block_attr(1)'运算符
				XMLDOM.SelectSingleNode("//template/if/block[@name='"&keys&"']/attr").SetAttribute "value", block_attr(2)'对比值
			Next			
			'获取block内容
			block_code = items(2)
			'创建block块
			'创建一个block节点,对于if块，这里就是if分支的代码,传递过来的是split()
			For I = 0 To Ubound(block_code)		
				Set blockNode = XMLDOM.CreateElement("tpl")
				blockNode.appendChild(XMLDOM.createCDATASection(block_code(I)))'给节点赋值,并且是用CDATA
				XMLDOM.SelectSingleNode("//template/if/block[@name='"&keys&"']").appendChild(blockNode)		
			Next
		Next
		
		'// 解析loop块到xml
		Dim arrAttr, oPI
		For Each keys in DIC_BLOCK_LOOP
			items = DIC_BLOCK_LOOP(keys)			
			'添加一个loopblock节点
			Set blockNode = XMLDOM.CreateElement("block")
			blockNode.SetAttribute "name", keys
			XMLDOM.SelectSingleNode("//template/loop").appendChild(blockNode)			
			'创建一个attr节点
			XMLDOM.SelectSingleNode("//template/loop/block[@name='"&keys&"']").appendChild(XMLDOM.CreateElement("attr"))			
			'获取loop 的attr属性	
			block_attr = items(1)
			'数组对象，0=名,1=值
			For I = 0 To Ubound(block_attr)
				'设置attr的属性
				XMLDOM.SelectSingleNode("//template/loop/block[@name='"&keys&"']/attr").SetAttribute block_attr(I, 0), block_attr(I, 1)'变量
			Next			
			'设置块
			'获取block内容
			block_code = items(2)
			'创建block块
			Set blockNode = XMLDOM.CreateElement("tpl")
			blockNode.appendChild(XMLDOM.createCDATASection(block_code))'给节点赋值,并且是用CDATA
			XMLDOM.SelectSingleNode("//template/loop/block[@name='"&keys&"']").appendChild(blockNode)		
		Next		
		'添加xml头部
		Set oPI = XMLDOM.CreateProcessingInstruction("xml", "version=""1.0"" encoding="""&strCharset&"""")
		XMLDOM.insertBefore oPI,XMLDOM.childNodes(0)		
		'保存到xml文件
		XMLDOM.save(path)		
		'销毁
		Set blockNode = Nothing
		Set tplNode = Nothing		
		Analysis_SaveToXml = System.Text.IIF((Err.Number = 0), True, False)
		Err.Clear
	End Function
	
	'// 缓存模板
	Private Sub FileCache()
		'// 自动创建缓存文件目录
		AutoCreateFolder(strSITEROOT & strCache_dir & strTemplate_dir)		
		'加载之后就解析模板
		call Analysis()		
		'思路，模板将缓存到三个文件，html,xml,asp html是解析后的模板代码,xml保存块参数,asp是带有块内容的可以通过include直接运行的
		'解析到xml
		call Analysis_SaveToXml(strCacheDir&"_.xml")
		'解析到html
		call SaveToFile(strTemplate, strCacheDir&"_.html", strCharset)
	End Sub

	'////////////////////////////////////////////////缓存部分///////////////////////////////////////////////////

	'// 缓存到内存
	Public Property Get addCache(ByRef Key, ByRef Content, ByRef cacheTime)
		'0=计数器/1=缓存时间/2=缓存内容/3=缓存有效期(有效期分钟)
		Dim Items(3): Items(0) = 0: Items(1) = Now(): Items(3) = CInt(cacheTime)'有效期		
		IF (IsObject(Content)) Then Set Items(2) = Content Else Items(2) = Content End If
		Application.Unlock
		Application(PREFIX & Key) = Items
		Application.Lock
	End Property
	
	'// 读取内存缓存取出变量 计数器++
	Public Property Get getCache(ByRef Key)
		Dim Items: Items = Application(PREFIX & Key)
		getCache = Empty		
		If (IsArray(Items)) Then
			'判断是否过期了
			If DateDiff("n", FormatDateTime(Items(1), 0), Now()) <= Items(3) Then
				If (IsObject(Items)) Then Set getCache = Items(2) Else getCache = Items(2) End If
				Application(PREFIX & Key)(0) = Application(PREFIX & Key)(0) + 1
			Else
				'清空这个缓存
				Application.Contents.Remove(PREFIX & Key)
			End If
		End If
	End Property
	
	'// 得到计数器数值，可以知道某个缓存被读取了多少次(也可以检测缓存是否存在,=-1 就是不存在)
	'// 如果缓存对象不存在就返回-1，如果存在就返回读取了多少次
	Public Property Get numCache(ByRef Key)
		Dim Items: Items = Application(PREFIX & Key)
		If (IsArray(Items)) Then
			numCache = Items(0)
			If DateDiff("n", FormatDateTime(Items(1), 0), Now()) > Items(3) Then
				Application.Contents.Remove(PREFIX & Key)
				numCache = -1
			End If
		Else numCache = -1 End If
	End Property
	
	'// 删除某个缓存
	Public Property Get clearOneClear(ByRef Key)
		Application.Contents.Remove(PREFIX & Key)
	End Property
	
	'// 清空所有缓存
	Public Property Get clearAllCache()
		Dim Key, Keys, KeyLength, KeyIndex
		For Each Key In Application.Contents
			If (Left(Key, Len(PREFIX)) = PREFIX) Then Keys = Keys & VBNewLine & Key
		Next		
		Keys = Split(Keys, vbNewLine)
		KeyLength = UBound(Keys)
		Application.Unlock
		For KeyIndex = 1 To KeyLength: Application.Contents.Remove(Keys(KeyIndex)): Next
		Application.Lock
	End Property

	'////////////////////////////////////////////////缓存部分///////////////////////////////////////////////////
	
	'// 该方法用来加载并且解析模板
	Public Property Get Load()
		'用不同的方式来加载模板
		If intCacheflag = 1 Then'使用文件缓存
			Dim FileCacheTimeOut: FileCacheTimeOut = False
			'// 如果有缓存就读取缓存，没有就加载模板后建立缓存
			If System.IO.ExistsFile(strCacheDir & "_.html") Then'如果能读取到缓存
				Dim Files : Set Files = FSO.GetFile(strCacheDir & "_.html")
				'// 检测文件缓存是否过期,如果没有超时
				If DateDiff("n", FormatDateTime(Files.DateLastModified, 0), Now()) < intCacheTplTime Then
					strTemplate = LoadFile( strCacheDir & "_.html" , strCharset)
				Else FileCacheTimeOut = True End If
			Else FileCacheTimeOut = True End If
			
			'如果读取不到缓存或者缓存过期,设置成不需要缓存,直接重新加载
			If FileCacheTimeOut Then
				intCacheflag = 0'设置不缓存
				call LoadTemplate()'加载模板
				call FileCache()'解析的块内容缓存到xml文件中,下次就不用在解析了
			End If
		ElseIf intCacheflag = 2 Then'内存缓存
			If Len(strTemplate_path) = 0 Then'如果不是文件模板
				Response.Write("内存缓存仅对文件模板有效")
			Else'开始读取内存
				strTemplate = getCache(strTemplate_path)
				If IsEmpty(strTemplate) Then'如果模板为空
					call LoadTemplate()'加载模板
					call addCache(strTemplate_path, strTemplate, intCacheTplTime)'缓存模板
				End If
			End If
		Else'0=没缓存
			call LoadTemplate()	'加载模板
			clearAllCache()		'清空所有模板内存缓存
		End If
		'加载之后就解析模板
		call Analysis()
	End Property
	
	'// 输出模板
	Public Property Get Display
		call AssignTpl()'解析数据 替换输出
		Response.Write(strTemplate)
		'如果设置了整页面缓存,就缓存这个页面
		If CBool(intCreateCachePage) Then call SaveToFile(strTemplate, strCachePage, strCharset)
	End Property
	
	'// 获得缓存路径
	Private Property Get strCacheDir
		strCacheDir = strSITEROOT & strCache_dir & strTemplate_dir & Replace(Replace(strTemplate_path, "/", "-"), "\", "-")
	End Property
	
	Private Property Get strMeta
		strMeta = "<meta name=""copyright"" content=""Boyle.ACL"">"
	End Property
	
	'// 清除所有模板缓存
	Public Property Get clearFileCache
		System.IO.DeleteFolder(strSITEROOT & strCache_dir)
	End Property
	
	'// 获取输出结果
	Public Property Get gethtml
		call AssignTpl()
		gethtml = strTemplate
	End Property
	
	'// 提取所有块标签
	Public Property Get getblock
		Set getblock = DIC_BLOCK_LOOP_LIST
	End Property
	
	'// 提取所有块标签的属性
	Public Property Get getattr
		Set getattr = DIC_BLOCK_LOOP_ATTR
	End Property
	
	'//---------------------------定义类的输入属性-------------------------------//
	'// 设置页面编码
	Public Property Let setCharset(ByVal strVar)
		strCharset = strVar
	End Property
	'// 设置网站目录
	Public Property Let setRoot(ByVal strVar)
		strSITEROOT = strVar
	End Property
	'// 设置模板目录
	Public Property Let setTemplatedir(ByVal strVar)
		strTemplate_dir = strVar
	End Property
	Public Property Let Root(ByVal strVar)
		strTemplate_dir = strVar
	End Property
	Public Property Get Root()
		Root = strTemplate_dir
	End Property
	'// 设置缓存目录
	Public Property Let setCachedir(ByVal strVar)
		strCache_dir = strVar
	End Property
	'// 设置模板文件路径，相对于模板目录
	Public Property Let setPath(ByVal strVar)
		strTemplate_path = strVar
	End Property
	'// 加载模板代码，如果不指定setPath,程序自动使用这里,使用这个之后模板缓存自动关闭
	Public Property Let setHtml(ByVal strVar)
		intCacheflag = 0'直接设置模板,缓存失效
		strTemplate = strVar
	End Property
	'// 设置缓存开关,默认是关闭
	Public Property Let setCacheType(ByVal intVar)
		intCacheflag = intVar
	End Property
	'// 设置缓存时间,单位是分钟,默认是10分钟
	Public Property Let setCacheTimeOut(ByVal strVar)
		intCacheTplTime = strVar
	End Property
	'// 设置缓存页面文件名
	Public Property Let setCachePageName(ByVal strVar)
		strCachePageName = strVar
	End Property
	'// 设置模板绝对路径,默认是开启
	'// 作用是输出的时候将模板相对路径替换成绝对路径,已经是绝对路径或者描点等不受影响
	Public Property Let setAbsPath(ByVal intVar)
		intAbsPath = intVar
	End Property

	'// 标签赋值
	Public Property Let add(ByVal strTag, ByVal strVal)
		Dim rs, I
		'单标签赋值
		strTag = LCase(strTag)
		If Left(strTag, 1) = "@" Then
			'如果是记录集赋值
			If TypeName(strVal) = "Recordset" Then
				Set rs = strVal
				'字段设置
				If rs.State And Not rs.Eof Then
					'将字段值自动赋值到标签中
					For I = 0 to rs.Fields.Count - 1
						strTag = "@" & LCase(rs.Fields(I).name)
						If DIC_BLOCK_ATTR.Exists(strTag) Then DIC_BLOCK_ATTR.Remove(strTag)
						DIC_BLOCK_ATTR(strTag) = System.Text.IIF(IsNull(rs(I)), "", rs(I))
					Next
					'rs.Close()'程序后面还可以引用rs记录集，所以这里不用关闭，不过在外面的时候记得要关闭
				End If
				Set rs = Nothing: Exit Property
			Else
				'否则就是一个一个的单标签赋值
				If DIC_BLOCK_ATTR.Exists(strTag) Then DIC_BLOCK_ATTR.Remove(strTag)
				DIC_BLOCK_ATTR(strTag) = System.Text.IIF(IsNull(strVal), "", strVal)
				Exit Property
			End If
		End If
		
		'循环体赋值,传递的是Rs
		If TypeName(strVal) = "Recordset" Then
			Set rs = strVal
			'字段设置
			If rs.State And Not rs.eof Then
				For I=0 to rs.Fields.count-1 
					RS_FIELD(strTag&"."&LCase(rs.Fields(I).name)) = I
				Next
				'赋值到循环块列表中
				If Not rs.Eof Then
					If DIC_BLOCK_LOOP_VAL.Exists(strTag) Then DIC_BLOCK_LOOP_VAL.Remove(strTag)
					DIC_BLOCK_LOOP_VAL(strTag) = rs.GetRows'赋值到循环块列表中
				End If
				'rs.Close()'程序后面还可以引用rs记录集，所以这里不用关闭，不过在外面的时候记得要关闭
			End If
			Set rs = Nothing
			Exit Property
		ElseIf TypeName(strVal) = "Variant()" Then'如果传递的是getrows
			If DIC_BLOCK_LOOP_VAL.Exists(strTag) Then DIC_BLOCK_LOOP_VAL.Remove(strTag)
		DIC_BLOCK_LOOP_VAL(strTag) = strVal
		ElseIf TypeName(strVal) = "String" Then'如果循环体直接赋值String，整个循环体当但标签处理
			If DIC_BLOCK_LOOP_VAL.Exists(strTag) Then DIC_BLOCK_LOOP_VAL.Remove(strTag)
			DIC_BLOCK_LOOP_VAL(strTag) = strVal
		End If
	End Property
	Public Property Let Assign(ByVal strTag, ByVal strVal)
		add(strTag) = strVal
	End Property

	'// 绑定字段
	Public Property Let bindField(ByVal strTag, ByVal strVal)
		If TypeName(strVal) = "Fields" Then
			Dim rs, I
			Set rs = strVal
			'字段设置
			For I = 0 to rs.Count - 1 
				RS_FIELD(strTag&"."&LCase(rs(I).name)) = I
			Next
			Set rs = Nothing
			Exit Property
		End If
	End Property
	
	Public Property Let CachePageTimeout(ByVal strVal)
		intCachePageTimeout = Int(strVal)
		If CBool(intCachePageTimeout) Then
			Dim CachePagePath: CachePagePath = "pagecache/"			
			call AutoCreateFolder(strSITEROOT & strCache_dir & CachePagePath)'自动生成存放目录			
			strCachePage = strSITEROOT & strCache_dir & CachePagePath & strCachePageName
			If System.IO.ExistsFile(strCachePage) Then'如果没有缓存
				Dim Files: Set Files = FSO.GetFile(strCachePage)
				'// 判断缓存是否超时
				If DateDiff("n", FormatDateTime(Files.DateLastModified, 0), Now()) < intCachePageTimeout Then
					'// 如果没有超时,读取缓存,至于缓存是用server.transfer还是fso,区别在于是否需要执行页面中的程序
					Server.Transfer(strCache_dir & CachePagePath & strCachePageName)
					'Response.Write(LoadFile(strCachePage, strCharset))
					Response.End()
				Else intCreateCachePage = True End If
			Else intCreateCachePage = True End If
		End If
	End Property
	
	'// 生成静态页面
	Public Property Let OutPutPage(ByVal strPath, ByVal strVar)
		Dim strOutputPath: strOutputPath = strPath & "/" & System.Text.IIF(Len(strVar), strVar, "index.html")
		If AutoCreateFolder(strPath) Then
			call AssignTpl()'数据替换输出
			call SaveToFile(strTemplate, strOutputPath, strCharset)
		End If
	End Property
	
	'// 设置页面并载入
	Public Sub File(ByVal strVar)
		strTemplate_path = strVar
		Load()
	End Sub
End Class
%>
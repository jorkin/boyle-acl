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
'// 作者：彭拉辉(kacarton@sohu.com)												//
'// 网址：http://blog.csdn.net/nhconch/archive/2004/07/10/38683.aspx			//
'// --------------------------------------------------------------------------- //

Class Cls_Template
	
	'// 定义私有命名对象
    Private m_Root, m_FileName, m_Unknowns
    Private m_ValueList, m_BlockList
    Private m_BlockMatches, m_ValueMatches
	
	'// 初始化资源	
    Private Sub Class_Initialize
        Set m_ValueList = Server.CreateObject("Scripting.Dictionary")
        Set m_BlockList = Server.CreateObject("Scripting.Dictionary")
        m_Root          = "."
        m_FileName      = ""
        m_Unknowns      = "remove"
		
		'// 初始化系统默认的错误信息
		System.Error.E(20001) = "目标模板文件为空或不存在，请检查路径是否正确。"
		System.Error.E(20002) = "未指定的区块标记。"
		System.Error.E(20003) = "未指定的文件标记。"
		System.Error.E(20004) = "未指定的模板标签名称。"
    End Sub
    
	'// 释放资源
    Private Sub Class_Terminate
        Set m_BlockMatches = Nothing
        Set m_ValueMatches = Nothing
    End Sub
    
	'/**
	' * @功能说明: 设置模板文件默认目录
	' * @参数说明: - blParam [string]: 目标文件夹路径
	' * @返回值:   - 
	' */
    Public Property Let Root(ByVal blParam)
        m_Root = blParam
    End Property
    Public Property Get Root()
        Root = m_Root
    End Property
    
	'/**
	' * @功能说明: 设置对未指定的标记的处理方式，有keep、remove、comment三种
	' * @返回值:   -
	' */
    Public Property Let Unknowns(ByVal unknown)
        m_Unknowns = unknown
    End Property
    Public Property Get Unknowns()
        Unknowns = m_Unknowns
    End Property
	
	'/**
	' * @功能说明: 设置模板文件(本类不支持多模板文件)
	' * @参数说明: - blHandle [string]: 保留参数
	' *  		   - blFile [string]: 目标文件名称
	' * @返回值:   - 
	' */
    Public Sub File(ByVal blHandle, ByVal blFile)
		m_FileName = blFile
		m_BlockList.Add blHandle, LoadFile()
	End Sub
	
	'/**
	' * @功能说明: 定义区块
	' * @参数说明: - blParent [string]: 上级区块名称
	' * 		  - blBlockTag [string]: 模板标签名称
	' * 		  - blName [bool]: 区块名称
	' */
    Public Sub Block(ByVal blParent, ByVal blBlockTag, ByVal blName)
        Dim Matches, Match
        System.Text.RegExpX.Pattern = "<!--\s+BEGIN " & blBlockTag & "\s+-->([\s\S.]*)<!--\s+END " & blBlockTag & "\s+-->"
        If Not m_BlockList.Exists(blParent) Then
			System.Error.Message = "("&blParent&")"
			System.Error.Raise 20002
		End  If
        Set Matches = System.Text.RegExpX.Execute(m_BlockList.Item(blParent))
        For Each Match In Matches
            m_BlockList.Add blBlockTag, Match.SubMatches(0)
            m_BlockList.Item(blParent) = Replace(m_BlockList.Item(blParent), Match.Value, "{" & blName & "}")
        Next
        Set Matches = Nothing
    End Sub
	
	'/**
	' * @功能说明: 移除同一个区块，用于区块循环输出
	' * @参数说明: - blName [string]: 区块名称
	' */
    Public Sub RemoveBlock(ByVal blName)
        If m_ValueList.Exists(blName) Then m_ValueList.Remove(blName)
    End Sub
	
	'/**
	' * @功能说明: 逐一对模板中相对应的标签进行映射
	' * @参数说明: - blName [string]: 模板标签名称
	' * 		  - blContent [string]: 映射的数据
	' * 		  - blAppend [bool]: 当出现相同标签，是否对内容进行追加。
	' */
    Public Sub Assign(ByVal blName, ByVal blContent, ByVal blAppend)
        Dim blContent1: blContent1 = System.Text.IIF(System.Text.IsEmptyAndNull(blContent), "", blContent)
        If m_ValueList.Exists(blName) Then
            If blAppend Then m_ValueList.Item(blName) = m_ValueList.Item(blName) & blContent1 _
            Else m_ValueList.Item(blName) = blContent1
        Else m_ValueList.Add blName, blContent1 End If
    End Sub
	
	'/**
	' * @功能说明: 批量对模板中相对应的标签进行映射
	' * @参数说明: - blArray [array]: 映射的数据。格式：[模板标签名称1:值1, 模板标签名称2:值2, ...]
	' * 		   - blAppend [bool]: 当出现相同标签，是否对内容进行追加。
	' */
	Public Sub AssignX(ByVal blArray, ByVal blAppend)
		Dim blDictionary: Set blDictionary = System.Text.ToHashTable(blArray)
		Dim blKey: blKey = blDictionary.Keys
		Dim blItem: blItem = blDictionary.Items
		Dim I: For I = 0 To blDictionary.Count - 1
			Assign blKey(I), blItem(I), blAppend
		Next
	End Sub
	
	'/**
	' * @功能说明: 解析区块
	' * @参数说明: - blName [string]: 区块名称
	' * 		  - blBlockTag [string]: 模板标签名称
	' * 		  - blAppend [bool]: 当出现相同标签，是否对内容进行追加，常常用于数据循环输出。
	' */
    Public Sub Parse(ByVal blName, ByVal blBlockTag, ByVal blAppend)
        If Not m_BlockList.Exists(blBlockTag) Then
			System.Error.Message = "("&blBlockTag&")"
			System.Error.Raise 20004
		End If
        If m_ValueList.Exists(blName) Then
            If blAppend Then m_ValueList.Item(blName) = m_ValueList.Item(blName) & InstanceValue(blBlockTag) _
            Else m_ValueList.Item(blName) = InstanceValue(blBlockTag)
        Else m_ValueList.Add blName, InstanceValue(blBlockTag) End If
    End Sub
	
	'/**
	' * @功能说明: 输出模板
	' * @参数说明: - blName [string]: 模板区块名称
	' */
    Public Sub Out(ByVal blName)
		Response.Write(Print(blName))
    End Sub
	
	'/**
	' * @功能说明: 保存为HTML文件
	' * @参数说明: - blFilePath [string]: 文件保存路径
	' *  		  - blName [string]: 模板区块名称
	' */
	Public Sub [Static](ByVal blFilePath, ByVal blName)
		System.IO.Save blFilePath, Print(blName)
	End Sub
    
	'// 读取文件内容
    Private Function LoadFile()
		With System.IO
			Dim blFilePath: blFilePath = .FormatFilePath(.Directory(m_Root, "") & "/" & m_FileName)
			LoadFile = .Read(blFilePath)
		End With
        If System.Text.IsEmptyAndNull(LoadFile) Then
			System.Error.Message = "("&blFilePath&")"
			System.Error.Raise 20001
		End If
    End Function
	
	'// 输出文件内容
	Private Function Print(ByVal blName)
		If Not m_ValueList.Exists(blName) Then
			System.Error.Message = "("&blName&")"
			System.Error.Raise 20003
		End If
		Print = Finish(m_ValueList.Item(blName))
	End Function
    
	'// 替换块内容
    Private Function InstanceValue(ByVal blBlockTag)
        InstanceValue = m_BlockList.Item(blBlockTag)
        Dim Keys: Keys = m_ValueList.Keys
        Dim I: For I = 0 To m_ValueList.Count - 1
			If System.Text.IsEmptyAndNull(m_ValueList.Item(Keys(I))) Then InstanceValue = Replace(InstanceValue, "{" & Keys(I) & "}", "") _
			Else InstanceValue = Replace(InstanceValue, "{" & Keys(I) & "}", m_ValueList.Item(Keys(I)))
        Next
    End Function
    
	'// 设定未指定映射的标记处理方式
    Private Function Finish(ByVal blParam1)
        Select Case UCase(m_Unknowns)
            Case "KEEP": Finish = blParam1
            Case "REMOVE": Finish = System.Text.ReplaceX("\{[^ \t\r\n}]+\}", blParam1, "")
            Case "COMMENT": Finish = System.Text.ReplaceX("\{([^ \t\r\n}]+)\}", blParam1, "<!-- Template Variable $1 undefined -->")
            Case Else Finish = blParam1
        End Select
    End Function
End Class
%>
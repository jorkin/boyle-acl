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
'// 2011/08/02 Boyle	 	 系统Cookie/Application操作类						//
'// --------------------------------------------------------------------------- //

Class Cls_IO_CAS
	
	'// 声明私有对象
	Private PrAES
	
	'/* 声明公共对象
	
	'// 初始化资源
	Private Sub Class_Initialize()
		PrAES = False
	End Sub
	
	'// 释放资源
	Private Sub Class_Terminate()
	End Sub
	
	Public Property Let AES(ByVal blParam)
		PrAES = System.Text.ToBoolean(blParam)
	End Property
	
	'// 获取一个Cookies值
	Public Function GetCookie(ByVal blParam)
		Dim blCookie, blName, blSubName
		If InStr(blParam, ":") > 0 Then
			blName = System.Text.CLeft(blParam, ":")
			blSubName = System.Text.CRight(blParam, ":")
			If Not System.Text.IsEmptyAndNull(blName) And Not System.Text.IsEmptyAndNull(blSubName) Then
				If Response.Cookies(blName).HasKeys Then blCookie = Request.Cookies(blName)(blSubName)
			End If
		Else
			If Not System.Text.IsEmptyAndNull(blParam) Then blCookie = Request.Cookies(blParam)
		End If
		If Not System.Text.IsEmptyAndNull(blCookie) Then
			If PrAES Then blCookie = System.Security.AES.Decrypt(blCookie)		
			GetCookie = blCookie
		Else GetCookie = "" End If
	End Function
	
	'// 设置一个Cookies值
	Public Sub SetCookie(ByVal blName, ByVal blValue, ByVal blConfig)
		Dim blExpires, blDomain, blPath, blSecure
		If isArray(blConfig) Then
			Dim I: For I = 0 To UBound(blConfig)
				If isDate(blConfig(I)) Then
					blExpires = cDate(blConfig(I))
				ElseIf System.Text.Test(blConfig(I), "INT") Then
					If blConfig(I) <> 0 Then blExpires = Now() + Int(blConfig(I)) / 60 / 24
				ElseIf System.Text.Test(blConfig(I), "DOMAIN") Or System.Text.Test(blConfig(I), "IP") Then
					blDomain = blConfig(I)
				ElseIf InStr(blConfig(I), "/") > 0 Then
					blPath = blConfig(I)
				ElseIf UCase(blConfig(I)) = "TRUE" Or UCase(blConfig(I)) = "FALSE" Then
					blSecure = blConfig(I)
				End If
			Next
		Else
			If isDate(blConfig) Then
				blExpires = cDate(blConfig)
			ElseIf System.Text.Test(blConfig, "INT") Then
				If blConfig <> 0 Then blExpires = Now() + Int(blConfig) / 60 / 24
			ElseIf System.Text.Test(blConfig, "DOMAIN") Or System.Text.Test(blConfig, "IP") Then
				blDomain = blConfig
			ElseIf InStr(blConfig, "/") > 0 Then
				blPath = blConfig
			ElseIf UCase(blConfig) = "TRUE" Or UCase(blConfig) = "FALSE" Then
				blSecure = blConfig
			End If
		End If
		If Not System.Text.IsEmptyAndNull(blValue) Then
			If PrAES Then blValue = System.Security.AES.Encrypt(blValue) End If
		End If
		If InStr(blName, ":") > 0 Then
			Dim blSubName: blSubName = System.Text.CRight(blName, ":")
			blName = System.Text.CLeft(blName, ":")
			Response.Cookies(blName)(blSubName) = blValue
		Else Response.Cookies(blName) = blValue End If
		If Not System.Text.IsEmptyAndNull(blExpires) Then Response.Cookies(blName).Expires = blExpires
		If Not System.Text.IsEmptyAndNull(blDomain) Then Response.Cookies(blName).Domain = blDomain
		If Not System.Text.IsEmptyAndNull(blPath) Then Response.Cookies(blName).Path = blPath
		If Not System.Text.IsEmptyAndNull(blSecure) Then Response.Cookies(blName).Secure = blSecure
	End Sub
	
	'// 删除一个Cookies值
	Public Sub RemoveCookie(ByVal blParam)
		Dim blName, blSubName
		If InStr(blParam, ":") > 0 Then
			blName = System.Text.CLeft(blParam,":")
			blSubName = System.Text.CRight(blParam, ":")
			If Not System.Text.IsEmptyAndNull(blName) And Not System.Text.IsEmptyAndNull(blSubName) Then
				If Response.Cookies(blName).HasKeys Then Response.Cookies(blName)(blSubName) = Empty
			End If
		Else
			If Not System.Text.IsEmptyAndNull(blParam) Then
				Response.Cookies(blParam) = Empty
				Response.Cookies(blParam).Expires = Now()
			End If
		End If
	End Sub
	
	'// 设置一个Application值
	Public Sub SetApplication(ByVal blName, ByRef blData)
		Application.Lock
		If IsObject(blData) Then Set Application(blName) = blData _
		Else Application(blName) = blData
		Application.UnLock
	End Sub
	
	'// 获取一个Application值
	Public Function GetApplication(ByVal blName)
		If Not System.Text.IsEmptyAndNull(blName) Then
			If IsObject(Application(blName)) Then Set GetApplication = Application(blName) _
			Else GetApplication = Application(blName)
		Else GetApplication = Empty End If
	End Function
	
	'// 删除一个Application值
	Public Sub RemoveApplication(ByVal blName)
		Application.Lock
		Application(blName) = Empty
		Application.UnLock
	End Sub	
	
End Class
%>
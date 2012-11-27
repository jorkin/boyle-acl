<!--#include file="./class_security_aes.asp"-->
<!--#include file="./class_security_md5.asp"-->
<!--#include file="./class_security_sha256.asp"-->
<%
'// --------------------------------------------------------------------------- //
'// Project Name		: Boyle.ACL												//
'// Program Name		: class_secrity.asp										//
'// Copyright Notice	: COPYRIGHT (C) 2011 BY BOYLE.							//
'// Creation Date		: 2011/08/02											//
'// Version				: 3.1.0.0802											//
'//																				//
'// Date       By			 Description										//
'// ---------- ------------- -------------------------------------------------- //
'// 2011/08/02 Boyle		 系统数据安全操作类									//
'// --------------------------------------------------------------------------- //

Class Cls_Security
	
	'// 定义私有命名对象
	Private Private_SHA256, Private_MD5, Private_AES
	
	'// 初始化类
	Private Sub Class_Initialize()
	End Sub	
	
	'// 释放类
	Private Sub Class_Terminate()
		If IsObject(Private_AES) Then Set Private_AES = Nothing
		If IsObject(Private_MD5) Then Set Private_MD5 = Nothing
		If IsObject(Private_SHA256) Then Set Private_SHA256 = Nothing
	End Sub
	
	'// 声明对象模块单元
	Public Property Get AES()
		If Not IsObject(Private_AES) Then Set Private_AES = New Cls_Security_AES End If
		Set AES = Private_AES
	End Property
	Public Property Get MD5(ByVal strVal, ByVal numVal)
		If Not IsOBject(Private_MD5) Then Set Private_MD5 = New Cls_Security_MD5 End If
		MD5 = Private_MD5.Encrypt(strVal, numVal)
	End Property
	Public Property Get SHA256(ByVal strVal)
		If Not IsObject(Private_SHA256) Then Set Private_SHA256 = New Cls_Security_SHA256 End If
		SHA256 = Private_SHA256.Encrypt(strVal)
	End Property
End Class
%>
<%
'// --------------------------------------------------------------------------- //
'// Project Name		: Boyle.ACL												//
'// Author				: Boyle(boyle7[at]qq.com)								//
'// Copyright Notice	: COPYRIGHT (C) 2011-2012 BY BOYLE.						//
'// Create Date			: 2012/11/28											//
'// Version				: 4.0.121028											//
'//																				//
'// Date       By			 Description										//
'// ---------- ------------- -------------------------------------------------- //
'// 2012/11/28 Boyle		 系统基础函数库										//
'// --------------------------------------------------------------------------- //

Public Sub A(byVal blParam)
	System.IO.Import LIB_PATH & "Action/"&blParam&"Action.class.asp"
End Sub

Public Sub M(byVal blParam)
	With System.Template
		.File blParam, blParam&"/"&blParam&".html"

		'.Assign "T_NAME", "Hello World!", False
		
		.Parse "OUT", blParam, False
		.Out   "OUT"
	End With
End Sub

Class Driver

	Public Name
	
	'// 初始化命名对象
	Private Sub Class_Initialize()

	End Sub
	
	'// 释放命名对象
	Private Sub Class_Terminate()
	End Sub
End Class
%>
<!--#include file="./Lib/Core/Boyle.class.asp"-->
<!--#include file="./Common/runtime.asp"-->
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
'// 2012/11/28 Boyle		 框架入口文件										//
'// --------------------------------------------------------------------------- //
%>

<%
'// 导入配置文件
System.IO.Import CONF_PATH & "config.asp"

'// 执行入口
System.Run()
'// 输出页面
'// http://localhost/?s=modle/action/var/value
'// http://localhost/app/index.php/Form/read/id/1

With System.Template
	'// 载入页面
	.Root = .Root & "/demo/"
	.File "index.html"

	.Assign("@title") = "模板示例 - Boyle.ACL"

	Dim Rs: Set Rs = System.Data.Query(System.Data.ToSQL(Array("BE_CUSTOMER", "ID,CI_NAME,CI_TELEPHONE,CI_ADDRESS", 5), "", ""))

	.Assign("customer") = Rs

	System.Data.C(Rs)

	Dim blPage: blPage = System.R("PAGE", 0)
	With System.Data
		.Page.CurrentPage = blPage
		.Page.PageSize = 15
		.Page.SQL = .ToSQL(Array("BE_CUSTOMER", "ID,CI_NAME,CI_TELEPHONE,CI_ADDRESS", ""), "", "")
		Dim blPageData: blPageData = .Page.Run
	End With

	.Assign("customerpage") = blPageData

	.Assign("@pager") = System.Data.Page.Out

	'// 输出页面
	.Display
End With

'A("Index")

Call Terminate()
%>
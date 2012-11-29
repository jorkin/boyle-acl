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

A("Index")

Call Terminate()
%>
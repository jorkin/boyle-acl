<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
'// --------------------------------------------------------------------------- //
'// Project Name		: Boyle.ACL												//
'// Author				: Boyle(boyle7[at]qq.com)								//
'// Copyright Notice	: COPYRIGHT (C) 2011-2012 BY BOYLE.						//
'// Create Date			: 2012/11/28											//
'// Version				: 4.0.121028											//
'//																				//
'// --------------------------------------------------------------------------- //
'// 项目入口文件																	//
'// --------------------------------------------------------------------------- //
OPTION EXPLICIT
Response.Buffer      = True
Server.ScriptTimeout = 90
Session.CodePage     = 65001
Session.LCID         = 2057
%>

<%
	'// 定义框架路径
	Private Const BOYLE_PATH = "./ACL/"
	'// 定义项目名称
	Private Const APP_NAME   = "DEMO"
	'// 定义项目路径
	Private Const APP_PATH   = "./Demo/"
%>
<!--#include file="./ACL/BoyleACL.asp"-->
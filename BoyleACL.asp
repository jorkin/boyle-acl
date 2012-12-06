<!--#include file="./Lib/Core/Boyle.class.asp"-->
<!--#include file="./Common/runtime.asp"-->
<%
'// +--------------------------------------------------------------------------
'// | Boyle.ACL [框架入口文件]
'// +--------------------------------------------------------------------------
'// | Copyright (c) 2008-2012 By Boyle. All rights reserved.
'// +--------------------------------------------------------------------------
'// | Licensed ( http://www.gnu.org/licenses/gpl.html )
'// +--------------------------------------------------------------------------
'// | Author: Boyle <boyle7[at]qq.com>
'// +--------------------------------------------------------------------------

'// 导入配置文件
System.IO.Import CONF_PATH & "config.asp"

'// 执行入口
System.Run()

'// 输出页面
'// http://localhost/?s=modle/action/var/value
'// http://localhost/app/index.php/Form/read/id/1

A("Index")
%>
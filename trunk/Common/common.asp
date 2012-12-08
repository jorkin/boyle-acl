<%
'// +--------------------------------------------------------------------------
'// | Boyle.ACL [系统基础函数库]
'// +--------------------------------------------------------------------------
'// | Copyright (c) 2008-2012 By Boyle. All rights reserved.
'// +--------------------------------------------------------------------------
'// | Licensed ( http://www.gnu.org/licenses/gpl.html )
'// +--------------------------------------------------------------------------
'// | Author: Boyle <boyle7[at]qq.com>
'// +--------------------------------------------------------------------------

'// 获取和设置配置参数
Public C: Set C = Dicary():C.CompareMode = 1

'// 释放数据源和基类
Public Sub Terminate()
	System.Data.DisConnect()
	Set System = Nothing
	Set C = Nothing
End Sub

Public Function M(byVal blParam)
	Set M = System.Model.New(C("DB.PREFIX") & blParam)
End Function
	
'// 创建一个字典对象
Public Function Dicary()
	Set Dicary = Server.CreateObject("Scripting.Dictionary")
End Function
%>
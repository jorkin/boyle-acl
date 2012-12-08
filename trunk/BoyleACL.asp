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

'// 读取配置信息
C("DB.PREFIX") = "BE_"

Dim Action, blUri: blUri = System.Get("s", 1)

If Not System.Text.IsEmptyAndNull(blUri) Then
	Dim blList: Set blList = System.Array.New
	blList.Symbol = "/"
	blList.Data = blUri
	Dim blModel, blAction, blVar, blValue
	Select Case blList.Size
		Case 1:
			blModel = blList(0): blAction = "Index": blVar = "p": blValue = "1"
		Case 2:
			blModel = blList(0): blAction = System.Text.IIF(System.Text.IsEmptyAndNull(blList(1)), "Index", blList(1))
			blVar = "p": blValue = "1"
		Case 3:
			blModel = blList(0): blAction = System.Text.IIF(System.Text.IsEmptyAndNull(blList(1)), "Index", blList(1))
			blVar = blList(2): blValue = "1"
		Case 4:
			blModel = blList(0): blAction = System.Text.IIF(System.Text.IsEmptyAndNull(blList(1)), "Index", blList(1))
			blVar = blList(2): blValue = blList(3)
	End Select
	blList(0) = blModel: blList(1) = blAction
	blList(2) = blVar: blList(3) = blValue

	'// 设置模板路径及文件
	System.Template.Root = System.Template.Root&"/"&blModel&"/"
	System.Template.File = blAction
	Show(blModel) '// 载入文件
	Set Action = Dicary()
	'On Error Resume Next
	Execute("Set Action("""&blModel&""") = New "&blModel&"Action")
	Execute("Action("""&blModel&""")."&blAction&"("""&blList.J(" ")&""")")
	'If Err Then Response.Redirect("./"): Err.Clear
	Set Action = Nothing: Set blList = Nothing
Else
	'// 设置模板路径及文件
	System.Template.Root = System.Template.Root&"/Index/"
	System.Template.File = "Index"
	Show("Index") '// 载入文件
	Set Action = New IndexAction
	Action.Index("Boyle.ACL")
	Set Action = Nothing
End If

Private Sub Show(ByVal blParam)
	System.IO.Import(LIB_PATH & "Action/"&blParam&"Action.class.asp")
End Sub

%>
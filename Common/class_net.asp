<%
'// --------------------------------------------------------------------------- //
'// Project Name		: Boyle.ACL												//
'// Program Name		: class_net.asp											//
'// Copyright Notice	: COPYRIGHT (C) 2011 BY BOYLE.							//
'// Creation Date		: 2011/08/02											//
'// Version				: 3.1.0.0802											//
'//																				//
'// Date       By			 Description										//
'// ---------- ------------- -------------------------------------------------- //
'// 2011/08/02 Boyle	 	 系统网络操作类										//
'// --------------------------------------------------------------------------- //

Class Cls_Net

	'// 声明公共对象
	
	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
	
	'// 功能说明: 获取客户端IP地址
	Public Function IP()
		Dim strIPAddr		
		If Request.ServerVariables("HTTP_X_FORWARDED_FOR") = "" Or InStr(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), "unknown") > 0 Then
			strIPAddr = Request.ServerVariables("REMOTE_ADDR")
		ElseIf InStr(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), ",") > 0 Then
			strIPAddr = Mid(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), 1, InStr(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), ",") - 1)
		ElseIf InStr(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), ";") > 0 Then
			strIPAddr = Mid(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), 1, InStr(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), ";") - 1)
		Else strIPAddr = Request.ServerVariables("HTTP_X_FORWARDED_FOR") End If
		IP = Trim(Mid(strIPAddr, 1, 30))
	End Function
	
	'/**
	' * @功能说明: 拒绝目标IP地址访问
	' * @参数说明: - blParam [string]: 被拒绝的目标IP
	' *			   - blRefuseList [array|string]: 非法IP列表，当为字符串时，不同项目之间用英文逗号隔开
	' * @返回值:   - [bool] 布尔值
	' */
	Public Function LockIP(ByVal blParam, ByVal blRefuseList)
		LockIP = False
		Dim oIP, tIP
		Dim strMatchs, strIP
		If Not IsArray(blRefuseList) Then blRefuseList = System.Text.ToArray(blRefuseList, ",")
		For Each oIP In blRefuseList
			oIP = Replace(oIP, "*", "\d*")
			tIP = sPlit(oIP, ".")
			System.Text.RegExpX.Pattern = "("&tIP(0)&"|)."&"("&tIP(1)&"|)."&"("&tIP(2)&"|)."&"("&tIP(3)&"|)"
			Set strMatchs = System.Text.RegExpX.Execute(blParam)
			strIP = strMatchs(0).SubMatches(0)&"."&strMatchs(0).SubMatches(1)&"."&strMatchs(0).SubMatches(2)&"."&strMatchs(0).SubMatches(3)
			If strIP = blParam Then LockIP = True: Exit Function
			Set strMatchs = Nothing
		Next
	End Function
	
	'// 功能说明: 将IP地址转换为数值
	Public Function IPToNumber(ByVal blParam)
		Dim blStrA, blStrB, blStrC, blStrD
		Dim blStrX: blStrX = blParam
		If IsNumeric(Left(blStrX, 2)) Then
			If blStrX = "127.0.0.1" Then blStrX = "192.168.0.1"
			blStrA = Left(blStrX, InStr(blStrX, ".") - 1)
			blStrX = Mid(blStrX,  InStr(blStrX, ".") + 1)
			blStrB = Left(blStrX, InStr(blStrX, ".") - 1)
			blStrX = Mid(blStrX,  InStr(blStrX, ".") + 1)
			blStrC = Left(blStrX, InStr(blStrX, ".") - 1)
			blStrD = Mid(blStrX,  InStr(blStrX, ".") + 1)
			If IsNumeric(blStrA) = 0 Or IsNumeric(blStrB) = 0 Or IsNumeric(blStrC) = 0 Or IsNumeric(blStrD) = 0 Then IPToNumber = 0 _
			Else IPToNumber = CLng(blStrA) * 256 * 256 * 256 + CLng(blStrB) * 256 * 256 + CLng(blStrC) * 256 + CLng(blStrD) - 1
		End If
	End Function

	'// 功能说明: 判断请求是否来自外部
	Public Function IsSelfPost()
		Dim HTTP_REFERER, SERVER_NAME
		HTTP_REFERER = CStr(Request.ServerVariables("HTTP_REFERER"))
		SERVER_NAME  = CStr(Request.ServerVariables("SERVER_NAME"))
		IsSelfPost = False
		IF Mid(HTTP_REFERER, 8, Len(SERVER_NAME)) = SERVER_NAME Then IsSelfPost = True
	End Function	
End Class
%>
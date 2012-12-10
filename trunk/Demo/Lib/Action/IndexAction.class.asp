<%
'// 本类由系统自动生成，仅供测试用途
Class IndexAction

	Private Sub Class_Initialize
	End Sub
	
	'// 释放资源
	Private Sub Class_Terminate()
		Call Terminate()
	End Sub

	'// 此方法为系统默认，请不要删除
	Public Sub Index()
		With System.Template
			.d("title") = "首页导航"
			.d("a") = "?m=index&a=parts"
			.d("b") = "?"&C("VAR_PATHINFO")&"=index"&C("URL_PATHINFO_DEPR")&"parts"
			.d("c") = U(Array("form", "add", "id"))
			
			.display()
		End With
	End Sub

	Public Sub Parts()
		With System.Template
			.d("title") = "Boyle.ACL 示例"
			
			Dim blPage
			'// 获取值，根据URL访问模式，自动获取值
			If C("URL_MODEL") = 0 Then blPage = System.R(":PAGE")
			If C("URL_MODEL") = 1 Then blPage = System.R(3)
			If C("URL_MODEL") = 2 Then blPage = ""

			Dim Parts: Set Parts = M("PARTS")
			Parts.Parameters("") = Array("CURRENTPAGE:"&blPage&"", "FIELD:ID,CP_NAME,CP_LOCALITY,CP_CAR")
			Dim PagerResult: PagerResult = Parts.Pager()
			.d("parts") = Array(PagerResult(0), "id,name,locality,car")
			.d("pager") = PagerResult(1)
			.d("sql") = PagerResult(2)("SQL")
			Set Parts = Nothing

			.display()
		End With
	End Sub
End Class
%>
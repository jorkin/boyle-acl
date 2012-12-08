<%
'// 本类由系统自动生成，仅供测试用途
Class IndexAction

	Private Sub Class_Initialize
		'System.Template.Root =  System.Template.Root&"/Index/"
	End Sub
	
	'// 释放资源
	Private Sub Class_Terminate()
		Call Terminate()
	End Sub

	'// 此方法为系统默认，请不要删除
	Public Sub Index(ByVal blParam)
		With System.Template
			.d("title") = blParam
			.Display()
		End With
	End Sub

	Public Sub Parts(ByVal blParam)
		With System.Template		
			.d("title") = "Boyle.ACL 示例"

			'// 获取数据
			blParam = System.Array.NewArray(blParam).Data
			'// 获取值，这里需要改进
			Dim blPage: blPage = blParam(3)'System.Get("PAGE", 0)
			Dim Parts: Set Parts = M("PARTS")
			Parts.Parameters("") = Array("CURRENTPAGE:"&blPage&"", "FIELD:ID,CP_NAME,CP_LOCALITY,CP_CAR", "URL:?s="&blParam(0)&"/"&blParam(1)&"/"&blParam(2)&"/*")
			Dim PagerResult: PagerResult = Parts.Pager()
			.d("parts") = Array(PagerResult(0), "id,name,locality,car")
			.d("pager") = PagerResult(1)
			.d("sql") = PagerResult(2)("SQL")
			Set Parts = Nothing

			.Display()
		End With
	End Sub

	'// 新建类实例
	Public Function [New]()
		Set [New] = New IndexAction
	End Function
End Class
%>
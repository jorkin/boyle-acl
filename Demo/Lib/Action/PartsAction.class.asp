<%
Class PartsAction

	Private Sub Class_Initialize
		System.Template.d("title") = "Boyle.ACL 示例 - 配件列表"
	End Sub
	
	'// 释放资源
	Private Sub Class_Terminate()
		Call Terminate()
	End Sub

	'// 此方法为系统默认，请不要删除
	Public Sub Index()
	End Sub

	Public Sub List()
		With System.Template
			Dim Data: Set Data = M("Parts")
			Data.Parameters("FIELD") = "ID,CP_SERIALNUMBER,CP_STOCKS,CP_NAME,CP_LOCALITY,CP_BRAND,CP_SIZE,CP_CAR,CP_WAREHOUSE,CP_UNIT,CP_SALEPRICE"
			Data.Parameters("") = Array("CURRENTPAGE:"&System.R(3)&"", "PAGESIZE:12")

			Dim PagerResult: PagerResult = Data.Pager()
			.d("vo") = Array(PagerResult(0), "id,sn,stocks,name,locality,brand,size,car,warehouse,unit,price")
			.d("pager") = PagerResult(1)
			.d("sql") = PagerResult(2)("SQL")
			Set Data = Nothing

			.display
		End With
	End Sub
End Class
%>
<%
Class FormAction

	Private Sub Class_Initialize
	End Sub
	
	'// 释放资源
	Private Sub Class_Terminate()
		Call Terminate()
	End Sub

	'// 此方法为系统默认，请不要删除
	Public Sub Index()
	End Sub

	Public Sub Save()
		With System.Template
			Dim blContent: blContent = System.Text.HtmlEncode(System.Post("blEditor"))
			.d("result") = blContent
			
			.display()
		End With
	End Sub

	Public Sub Delete()
		With System.Template
			.d("result") = "msg"
			
			.display()
		End With
	End Sub

	Public Sub Add()
		With System.Template
			.d("title") = "Boyle.ACL 示例 - 表单示例"

			If C("URL_MODEL") = 0 Then .d("posturl") = "?m=form&a=save"
			If C("URL_MODEL") = 1 Then .d("posturl") = "?s=form/save"

			.display()
		End With
	End Sub

	Public Sub List()
		With System.Template
			.d("title") = "Boyle.ACL 示例 - 表单示例"

			Dim Customer: Set Customer = M("Customer")
			Customer.Parameters("FIELD") = "ID,CI_SERIALNUMBER,CI_NAME"
			Customer.Parameters("WHERE") = "ID=1"
			Customer.Parameters("") = Array("CURRENTPAGE:"&System.R(3)&"", "PAGESIZE:15")

			Dim PagerResult: PagerResult = Customer.Pager()
			.d("customer") = Array(PagerResult(0), "id,sn,name")
			.d("pager") = PagerResult(1)
			.d("sql") = PagerResult(2)("SQL")
			Set Customer = Nothing

			.display()
		End With
	End Sub

	Public Sub Edit()
		With System.Template
			.d("title") = "Boyle.ACL 示例 - 表单示例"

			Dim Customer: Set Customer = M("Customer")
			Customer.Parameters("") = "WHERE:ID="&System.R(3)
			Customer.Parameters("FIELD") = "ID,CI_SERIALNUMBER,CI_NAME,CI_ADDRESS,CI_TELEPHONE,CI_FAX,CI_CELLPHONE,CI_REMARK AS CONTENT"
			Dim blData: Set blData = Customer.Select()
			Dim blId: blId = blData(0)
			.d("$") = Array(blData, "ID")
			.d("sql") = Customer.Parameters("SQL")
			Set Customer = Nothing


			If C("URL_MODEL") = 0 Then .d("posturl") = "?m=form&a=save&id="&blId
			If C("URL_MODEL") = 1 Then .d("posturl") = "?s=form/save/id/"&blId

			.display()
		End With
	End Sub
End Class
%>
<%
C("Validate") = Array(	Array("content", "require", "内容不能为空"), _
						Array("ci_name1", "require", "名称不能为空"), _
						Array("ci_fax1", "phone", "传真号码格式有误"), _
						Array("ci_telephone1", "phone", "电话号码格式有误"), _
						Array("ci_cellphone", "mobile", "手机号码格式有误"))

C("Auto") = Array(Array("ci_createtime", "time", 1, "function"))
%>
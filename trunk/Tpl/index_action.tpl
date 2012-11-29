<%
'// 本类由系统自动生成，仅供测试用途
With System.Template
	.File "INDEX", "index/index.html"

	.Assign "T_NAME", "Hello World!", False
	
	.Parse "OUT", "INDEX", False
	.Out   "OUT"
End With
%>
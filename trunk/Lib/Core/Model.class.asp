<%
'// +----------------------------------------------------------------------
'// | Boyle.ACL [系统模型操作类]
'// +----------------------------------------------------------------------
'// | Copyright (c) 2008-2012 By Boyle. All rights reserved.
'// +----------------------------------------------------------------------
'// | Licensed ( http://www.gnu.org/licenses/gpl.html )
'// +----------------------------------------------------------------------
'// | Author: Boyle <boyle7[at]qq.com>
'// +----------------------------------------------------------------------

Class Cls_Model

	'// 定义私有命名对象
	Private PrDic, PrTable
	Private PrError

	'// 定义公共命名对象
	Public Validate, Auto

	Private Sub Class_Initialize
		Set PrDic = Dicary(): PrDic.CompareMode = 1
		Set PrError = Dicary(): PrError.CompareMode = 1
		
		PrTable = "" '// 初始化表格名称
	End Sub
	Private Sub Class_Terminate
		Set PrDic = Nothing
		Set PrError = Nothing
	End Sub
	
	'// 新建类实例
	Public Function [New](ByVal bParam)
		Set [New] = New Cls_Model
		[New].Table = bParam
	End Function

	'// 设置读取表格名称
	Public Property Get Table()
		Table = PrTable
	End Property
	Public Property Let Table(ByVal bParam)
		PrTable = bParam
	End Property

	'// 实现批量设置SQL语句参数
	Public Property Let Parameters(ByVal bField, ByVal bValue)
		'// 当键值为空时，表示对所有参数进行设置
		If System.Text.IsEmptyAndNull(bField) Then
			Dim tmpDic, tmpKey
			Select Case VarType(bValue)
				Case 0, 1: '// vbEmpty,vbNull
					Set tmpDic = Dicary(): PrDic.RemoveAll'// 清空所有配置参数
				Case 2, 3, 4, 5: '// vbInteger,vbLong,vbSingle,vbDouble
					Set tmpDic = System.Text.ToHashTable(Array("LIMIT:"&bValue))
				'Case 6: '// vbCurrency
				'Case 7: '// vbDate
				Case 8: '// vbString
					'// 如果目标参数的值为字符串时，将其转换为数组
					'// 其中对字符串用“|”符号进行分隔
					Dim tmpObj: Set tmpObj = System.Array.New
					tmpObj.Symbol = "|"
					tmpObj.Data = bValue
					Set tmpDic = System.Text.ToHashTable(tmpObj.ToArray)
					Set tmpObj = Nothing
				Case 9: '// vbObject
					Set tmpDic = bValue
				'Case 10: '// vbError
				'Case 11: '// vbBoolean
				Case 8192, 8194, 8204, 8209: '// 8192(Array),8204(vbVariant()),8209(Byte)
					Set tmpDic = System.Text.ToHashTable(bValue)
			End Select
			For Each tmpKey In tmpDic: PrDic(tmpKey) = tmpDic.Item(tmpKey): Next
			Set tmpDic = Nothing
		Else
			Select Case VarType(bValue)
				Case 0, 1:
					PrDic(bField) = ""
				'Case 2, 3, 4, 5, 6, 7, 8, 11:
				''	PrDic(bField) = bValue
				'Case 9:
				Case 8192, 8194, 8204, 8209:
					If UCase(bField) = "WHERE" Then
						PrDic(bField) = JoinWhere(bValue)
					ElseIf UCase(bField) = "FIELD" Then
						PrDic(bField) = System.Array.NewArray(bValue).J(",")
					Else PrDic(bField) = bValue(0) End If
				Case Else
					PrDic(bField) = bValue
			End Select
		End If
	End Property
	
	'// 获取参数集合
	'// 如果参数为空时，则返回一个DIC对象，否则返回目标项的值
	Public Property Get Parameters(ByVal bField)
		If System.Text.IsEmptyAndNull(bField) Then Set Parameters = PrDic _
		Else Parameters = PrDic(bField)
	End Property

	'// 拼接条件语句
	Private Function JoinWhere(ByVal bValue)
		Dim tmpObj: Set tmpObj = System.Array.NewHash(bValue)
		'// 判断是否存在逻辑判断符，如果存在则用此符号进行组装，否则用AND进行组装
		If LCase(tmpObj.HasIndex("_logic")) Then
			'// 获取逻辑判断符的值
			Dim blLogic: blLogic = tmpObj("_logic")
			'// 删除逻辑判断符 记录项
			tmpObj.Delete("_logic")
			JoinWhere = tmpObj.J(") "& blLogic &" (")
		Else JoinWhere = tmpObj.J(") AND (") End If
		Set tmpObj = Nothing
	End Function

	'// 创建数据对象 但不保存到数据库
	Public Function Create(ByVal bValue)
		'// 自动完成
		'// 完成字段，完成规则，[完成条件，附加规则]
		'// 完成条件：（可选）包括：1:新增数据的时候处理（默认）/2:更新数据的时候处理/3:所有情况都进行处理


		'// 自动验证
		'// 验证字段，验证规则，错误提示，[验证条件，附加规则，验证时间]
		'// 验证条件：（可选）包含下面几种情况：0:存在字段就验证（默认）/1:必须验证/2:值不为空的时候验证
		'// 附加规则：（可选）配合验证规则使用，包括下面一些规则：
		'// 验证时间：（可选） 1:新增数据时候验证/2:编辑数据时候验证/3:全部情况下验证（默认）
		Dim blErrDic: Set blErrDic = Dicary()
		If Not System.Text.IsEmptyAndNull(Validate) Then
			Dim I, J, blValiArr
			Dim blValiDic: Set blValiDic = Dicary()
			'// 获取验证规则
			For I = 0 To UBound(Validate)
				Set blValiArr = System.Array.NewArray(Validate(I))
				'// 对验证规则进行参数补全
				If blValiArr.Size < 4 Then blValiArr.Insert 3, Array(0, "regex", 3)
				If blValiArr.Size < 5 Then blValiArr.Insert 4, Array("regex", 3)
				If blValiArr.Size < 6 Then blValiArr.Insert 5, Array(3)
				'// 将验证规则保存在字典中
				Set blValiDic(LCase(blValiArr(0))) = blValiArr
			Next
			'// 获取字段列表
			Dim blName, blContent
			Dim blRule, blAttach, blMessage
			'Dim blErrList: ReDim blErrList(blValiDic.Count-1, 1)
			For I = 0 To UBound(bValue)
				blName = LCase(bValue(I)(0))
				blContent = bValue(I)(1)
				If blValiDic.Exists(blName) Then
					'// 取得[验证规则]的值
					blRule = blValiDic(blName)(1)
					'// 取得[错误提示]的值
					blMessage = blValiDic(blName)(2)
					'// 取得[附加规则]的值
					blAttach = blValiDic(blName)(4)
					Select Case LCase(blAttach)
						'// 正则验证，定义的验证规则是一个正则表达式（默认）
						Case "regex"
							'// 如果存在一个规则不通过，则立即退出循环
							If Not System.Text.Test(blContent, blRule) Then								
								blErrDic("title") = blName
								blErrDic("message") = blMessage
								Exit For
								'// 以下是用二维数组来保存多条错误信息，暂时不采用
								'blErrList(I, 0) = blName
								'blErrList(I, 1) = blMessage
							End If
						'// 函数验证，定义的验证规则是一个函数名
						'Case "function"
						'// callback	方法验证，定义的验证规则是当前模型类的一个方法
						'// confirm		验证表单中的两个字段是否相同，定义的验证规则是一个字段名
						'// equal		验证是否等于某个值，该值由前面的验证规则定义
						'// notequal	验证是否不等于某个值，该值由前面的验证规则定义（3.1.2版本新增）
						'// in			验证是否在某个范围内，定义的验证规则可以是一个数组或者逗号分割的字符串
						'// notin		验证是否不在某个范围内，定义的验证规则可以是一个数组或者逗号分割的字符串
						'// length		验证长度，定义的验证规则可以是一个数字（表示固定长度）或者数字范围（例如3,12 表示长度从3到12的范围）
						'// between		验证范围，定义的验证规则表示范围，可以使用字符串或者数组，例如1,31或者array(1,31)
						'// notbetween	验证不在某个范围，定义的验证规则表示范围，可以使用字符串或者数组
						'// expire		验证是否在有效期，定义的验证规则表示时间范围，可以到时间，例如可以使用 2012-1-15,2013-1-15 表示当前提交有效期在2012-1-15到2013-1-15之间，也可以使用时间戳定义
						'// ip_allow	验证IP是否允许，定义的验证规则表示允许的IP地址列表，用逗号分隔，例如201.12.2.5,201.12.2.6
						'// ip_deny		验证IP是否禁止，定义的验证规则表示禁止的ip地址列表，用逗号分隔，例如201.12.2.5,201.12.2.6
						'// unique		验证是否唯一，系统会根据字段目前的值查询数据库来判断是否存在相同的值。
					End Select
				End If
			Next
			Set blValiDic = Nothing

			'// 判断是否存在验证未通过项
			If blErrDic.Count > 0 Then
				Set PrError = blErrDic: Set blErrDic = Nothing: Create = False
			Else Create = True End If
		Else Create = True End If
	End Function

	'// 返回错误信息
	Public Function GetError()
		If PrError.Count > 0 Then
			Dim ErrTpl: Set ErrTpl = System.Template.New(BOYLE_PATH&"/Tpl/dispatch_jump.tpl")
			ErrTpl.d("block") = PrError
			GetError = ErrTpl.GetHtml()
			Set ErrTpl = Nothing
		End If
	End Function

	'// 新增数据
	Public Function Add(ByVal bValue)
		With System.Data
			PrDic("SQL") = .ToSQL(Array(PrTable, PrDic("FIELD"), PrDic("LIMIT")), PrDic("WHERE"), PrDic("ORDER"))
			Add = .Create(PrDic("SQL"), bValue)
		End With
	End Function

	'// 保存数据
	Public Function Save(ByVal bValue)
		With System.Data
			PrDic("SQL") = .ToSQL(Array(PrTable, PrDic("FIELD"), PrDic("LIMIT")), PrDic("WHERE"), PrDic("ORDER"))
			Save = .Update(PrDic("SQL"), bValue)
		End With
	End Function

	'// 查询数据
	Public Function [Select]()
		With System.Data
			PrDic("SQL") = .ToSQL(Array(PrTable, PrDic("FIELD"), PrDic("LIMIT")), PrDic("WHERE"), PrDic("ORDER"))
			Set [Select] = .Read(PrDic("SQL"))
		End With
	End Function

	'// 查询数据并分页
	Public Function Pager()
		With System.Data
			PrDic("SQL") = .ToSQL(Array(PrTable, PrDic("FIELD"), PrDic("LIMIT")), PrDic("WHERE"), PrDic("ORDER"))
			'// 将所有参数传递给分页类
			.Page.Parameters("") = Me.Parameters("")
			'// 对得到的结果进行行列对换
			Dim blList: blList = System.Array.Swap(.Page.Run)
			'// 返回数组，顺序依次为 [0]记录集列表，[1]分页导航码，[2]分页参数
			Pager = Array(blList, .Page.Out, .Page.Parameters(""))
		End With
	End Function

	'// 删除数据
	Public Function Delete(ByVal bValue)
		If Not System.Text.IsEmptyAndNull(bValue) Then PrDic("WHERE") = bValue
		PrDic("SQL") = System.Data.ToSQL(Array(PrTable, PrDic("FIELD"), PrDic("LIMIT")), PrDic("WHERE"), PrDic("ORDER"))
		Delete = System.Data.Delete(PrDic("SQL"))
	End Function

	'// 统计查询
	'// 统计数量，参数是统计的字段名（可选）
	Public Function Count(ByVal bValue)
		With System.Text
			Dim blField: blField = .IIF(Not .IsEmptyAndNull(bValue), bValue, "*")
			Dim blSQL: blSQL = "Select Count("&blField&") From "&PrTable&""
			PrDic("SQL") = .IIF(Not .IsEmptyAndNull(PrDic("WHERE")), (blSQL & " Where (" & PrDic("WHERE")) & ")", blSQL)
		End With
		Count = System.Data.Read(PrDic("SQL"))(0)
	End Function
	'// 获取最大值，参数是要统计的字段名（必须）
	Public Function Max(ByVal bValue)
		With System.Text
			Dim blSQL: blSQL = "Select Max("&bValue&") From "&PrTable&""
			PrDic("SQL") = .IIF(Not .IsEmptyAndNull(PrDic("WHERE")), (blSQL & " Where (" & PrDic("WHERE")) & ")", blSQL)
		End With
		Max = System.Data.Read(PrDic("SQL"))(0)
	End Function
	'// 获取最小值，参数是要统计的字段名（必须）
	Public Function Min(ByVal bValue)
		With System.Text
			Dim blSQL: blSQL = "Select Min("&bValue&") From "&PrTable&""
			PrDic("SQL") = .IIF(Not .IsEmptyAndNull(PrDic("WHERE")), (blSQL & " Where (" & PrDic("WHERE")) & ")", blSQL)
		End With
		Min = System.Data.Read(PrDic("SQL"))(0)
	End Function
	'// 获取平均值，参数是要统计的字段名（必须）
	Public Function Avg(ByVal bValue)
		With System.Text
			Dim blSQL: blSQL = "Select Avg("&bValue&") From "&PrTable&""
			PrDic("SQL") = .IIF(Not .IsEmptyAndNull(PrDic("WHERE")), (blSQL & " Where (" & PrDic("WHERE")) & ")", blSQL)
		End With
		Avg = System.Data.Read(PrDic("SQL"))(0)
	End Function
	'// 获取总分，参数是要统计的字段名（必须）
	Public Function Sum(ByVal bValue)
		With System.Text
			Dim blSQL: blSQL = "Select Sum("&bValue&") From "&PrTable&""
			PrDic("SQL") = .IIF(Not .IsEmptyAndNull(PrDic("WHERE")), (blSQL & " Where (" & PrDic("WHERE")) & ")", blSQL)
		End With
		Sum = System.Data.Read(PrDic("SQL"))(0)
	End Function

	'// 字段值增长，只对单条记录进行更改
	'// bValue[:step]
	Public Function setInc(ByVal bValue)
		With System.Text
			Dim blField: blField = .Separate(bValue)
			Dim blStep: blStep = .IIF(Not .IsEmptyAndNull(blField(1)), blField(1), 1)
			Dim blSQL: blSQL = "Select Top 1 "&blField(0)&" From "&PrTable&""
			PrDic("SQL") = .IIF(Not .IsEmptyAndNull(PrDic("WHERE")), (blSQL & " Where (" & PrDic("WHERE")) & ")", blSQL)
			Dim blSourceValue: blSourceValue = System.Data.Read(PrDic("SQL"))(0)
			blSourceValue = .IIF(IsNumeric(blSourceValue), blSourceValue, 0)
		End With
		setInc = System.Data.Update(PrDic("SQL"), Array(Array(blField(0), blSourceValue + blStep)))
	End Function

	'// 字段值减少，只对单条记录进行更改
	'// bValue[:step]
	Public Function setDec(ByVal bValue)
		With System.Text
			Dim blField: blField = .Separate(bValue)
			Dim blStep: blStep = .IIF(Not .IsEmptyAndNull(blField(1)), blField(1), 1)
			Dim blSQL: blSQL = "Select Top 1 "&blField(0)&" From "&PrTable&""
			PrDic("SQL") = .IIF(Not .IsEmptyAndNull(PrDic("WHERE")), (blSQL & " Where (" & PrDic("WHERE")) & ")", blSQL)
			Dim blSourceValue: blSourceValue = System.Data.Read(PrDic("SQL"))(0)
			blSourceValue = .IIF(IsNumeric(blSourceValue), blSourceValue, 0)
		End With
		setDec = System.Data.Update(PrDic("SQL"), Array(Array(blField(0), blSourceValue - blStep)))
	End Function
End Class
%>
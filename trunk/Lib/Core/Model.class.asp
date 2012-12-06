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
	Private PrDic
	Private PrTable, PrField, PrLimit, PrWhere, PrOrder

	Private Sub Class_Initialize
		Set PrDic = Dicary(): PrDic.CompareMode = 1
		
		PrTable = "" '// 初始化表格名称
	End Sub
	Private Sub Class_Terminate
		Set PrDic = Nothing
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
		PrTable = C("DB.PREFIX") & bParam
	End Property

	'// 设置读取表格字段名称
	Public Property Get Field()
		Field = PrField
	End Property
	Public Property Let Field(ByVal bParam)
		PrField = bParam
	End Property

	'// 设置读取的记录数
	Public Property Get Limit()
		Limit = PrLimit
	End Property
	Public Property Let Limit(ByVal bParam)
		PrLimit = bParam
	End Property

	'// 设置读取查询条件
	Public Property Get Where()
		Where = PrWhere
	End Property
	Public Property Let Where(ByVal bParam)
		PrWhere = bParam
	End Property

	'// 设置读取查询条件
	Public Property Get Order()
		Order = PrOrder
	End Property
	Public Property Let Order(ByVal bParam)
		PrOrder = bParam
	End Property

	'// 实现批量设置SQL语句参数
	Public Property Let Parameter(ByVal bField, ByVal bValue)
		If Not System.Text.IsEmptyAndNull(bField) Then
			Select Case VarType(bValue)
				Case 0, 1, 8: '// vbNull,vbEmpty,vbString
					PrDic(bField) = ""&bValue
				Case 2, 3, 4, 5: '//vbInteger,vbLong,vbSingle,vbDouble
					PrDic(bField) = bValue
				Case 9: '// vbObject
					Select Case TypeName(bValue)
						Case "Cls_Array":
							'// 对条件进行组装
							If UCase(bField) = "WHERE" Then
								'PrDic(bField) = JoinWhere(bValue)
							End If
					End Select
				Case 8192, 8194, 8204, 8209: '// 8192(vbArray),8204(vbVariant),8209(vbByte)
					Dim blStr
					If UCase(bField) = "WHERE" Then
						Dim blLogic: blLogic = System.Text.IIF(UBound(bValue), System.Array.NewHash(bValue(1))("_logic"), "AND")
						blStr = System.Array.NewArray(bValue(0)).J(" "& blLogic &" ")
					End If
					PrDic(bField) = blStr
			End Select
		End If
	End Property

	'// 拼接条件语句
	Private Function JoinWhere(ByVal bValue)
		'// 判断是否存在逻辑判断符，如果存在则用此符号进行组装，否则用AND进行组装
		If LCase(bValue.HasIndex("_logic")) Then
			'// 获取逻辑判断符
			Dim blLogic: blLogic = bValue("_logic")
			'// 删除逻辑判断符 记录项
			bValue.Delete("_logic")
			Dim blStr, blList: blList = bValue.Hash
			Dim I: For I = 0 To UBound(blList)
				blStr = blStr & " " & blLogic & " " & blList(I)
			Next
			JoinWhere = blStr
		Else
			'// 删除逻辑判断符项
			bValue.Delete("_logic")
			JoinWhere = bValue.J(" AND ")
		End If	
	End Function

	'// 新增数据
	Public Sub Add()
	End Sub

	'// 保存数据
	Public Sub Save()
	End Sub

	'// 查询数据
	Public Function [Select]()
		With System.Data
			Dim blSQL: blSQL = .ToSQL(Array(PrTable, PrDic("FIELD"), PrDic("LIMIT")), PrDic("WHERE"), PrDic("ORDER"))
			System.WE blSQL
			System.Template.d("sql") = blSQL
			Set [Select] = .Read(blSQL)
		End With
	End Function

	'// 删除数据
	Public Sub Delete()
	End Sub

	'// 设置记录的某个字段值
	Public Function setField()
	End Function

	'// 获取一条记录的某个字段值
	Public Function getField()
	End Function

	'// 字段值增长
	Public Function setInc()
	End Function

	'// 字段值减少
	Public Function setDec()
	End Function
End Class
%>
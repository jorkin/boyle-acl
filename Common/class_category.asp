<%
'// --------------------------------------------------------------------------- //
'// Project Name		: Boyle.ACL												//
'// Program Name		: class_category.asp									//
'// Copyright Notice	: COPYRIGHT (C) 2011 BY BOYLE.							//
'// Creation Date		: 2011/08/02											//
'// Version				: 3.1.0.0802											//
'//																				//
'// Date       By			 Description										//
'// ---------- ------------- -------------------------------------------------- //
'// 2011/08/02 Boyle		 系统分类操作类										//
'// --------------------------------------------------------------------------- //

Class Cls_Category
	
	'// 声明私有对象
	Private prData
	Private I, N, K		                        '// N为分类级别, I为数组标量
	Private prMaxArray, prArrayReturn()         '// prArrayReturn为排序后返回数组
	
    Private Sub Class_Initialize()
		I = 0
    End Sub
	Private Sub Class_Terminate()
	End Sub
	
	'// 设置分类数据
	Public Property Let Data(byVal blData)
		prData = blData
	End Property
	
	'// 取得下级分类，注意blCount参数是必须的
	'// 因为递归调用的原因，所以每次执行函数时blCount这个变量都要以不同的名字出现计数
	Private Function PreArray(ByVal blPreviousId, ByVal blCount)
		'// 将子类层数号叠加
		N = N + 1
		For blCount = 0 To prMaxArray
			'// 只要存在子类就进行输出
			If prData(1, blCount) = blPreviousId Then
				prArrayReturn(0, I) = prData(0, blCount)
				prArrayReturn(1, I) = N
				prArrayReturn(2, I) = prData(2, blCount)
				I = I + 1
				
				'// 递归取得父类下的所有子类
				Call PreArray(prData(0, blCount), "COUNT"&prData(0, blCount))
			End If
		Next
		'// 因为递归调用，必须叠加之后再还原
		N = N - 1
	End Function
	
	'// 重新对分类数据进行排序，让子类跟随着父类
	'// 取得所有根类，调用PreArray取得下级分类，得到所有
	Public Function Sort()
		prMaxArray = UBound(prData, 2)
		'// 声明一组二维数组
		reDim prArrayReturn(2, prMaxArray)
		For K = 0 To prMaxArray
			'// 只输出根类
			If prData(1, K) = 0 Then
				N = 0
				'// 唯一识别号(ID)
				prArrayReturn(0, I) = prData(0, K)
				'// 将父类识别号，重置为子类层数号（根类为0）
				prArrayReturn(1, I) = N
				'// 分类名称
				prArrayReturn(2, I) = prData(2, K)
				I = I + 1
				
				'// 递归取得此根类下所有子类
				Call PreArray(prData(0, K), "COUNT"&prData(0, K))
			End If
		Next: I = 0 '// 取出所有分类之后，复位I值
		Sort = prArrayReturn
	End Function
	
End Class
%>
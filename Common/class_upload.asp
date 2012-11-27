<%
'// --------------------------------------------------------------------------- //
'// Project Name		: Boyle.ACL												//
'// Program Name		: class_upload.asp										//
'// Copyright Notice	: COPYRIGHT (C) 2011 BY BOYLE.							//
'// Creation Date		: 2011/08/02											//
'// Version				: 3.1.0.0802											//
'//																				//
'// Date       By			 Description										//
'// ---------- ------------- -------------------------------------------------- //
'// 2011/08/02 Boyle	 	 系统文件上传操作类									//
'// --------------------------------------------------------------------------- //

'// --------------------------------------------------------------------------- //
'// 作者：风声(风声无组件上传类 2.11)											//
'// 网坦：http://www.fonshen.com												//
'// --------------------------------------------------------------------------- //

Class Cls_Upload

	Private m_TotalSize, m_MaxSize, m_FileType, m_SavePath, m_AutoSave, m_Error, m_Charset
	Private m_dicForm, m_binForm, m_binItem, m_strDate, m_lngTime
	Public	FormItem, FileItem

	Public Property Get Version
		Version = "Fonshen ASP UpLoadClass Version 2.11"
	End Property

	Public Property Get Error
		Error = m_Error
	End Property

	Public Property Get Charset
		Charset = m_Charset
	End Property
	Public Property Let Charset(strCharset)
		m_Charset = strCharset
	End Property

	Public Property Get TotalSize
		TotalSize = m_TotalSize
	End Property
	Public Property Let TotalSize(lngSize)
		If isNumeric(lngSize) Then m_TotalSize = Clng(lngSize)
	End Property

	Public Property Get MaxSize
		MaxSize = m_MaxSize
	End Property
	Public Property Let MaxSize(lngSize)
		If isNumeric(lngSize) Then m_MaxSize = Clng(lngSize)
	End Property

	Public Property Get FileType
		FileType = m_FileType
	End Property
	Public Property Let FileType(strType)
		m_FileType = strType
	End Property

	Public Property Get SavePath
		SavePath = m_SavePath
	End Property
	Public Property Let SavePath(strPath)
		m_SavePath = Replace(strPath, Chr(0), "")
	End Property

	Public Property Get AutoSave
		AutoSave = m_AutoSave
	End Property
	Public Property Let AutoSave(byVal Flag)
		Select Case Flag
			Case 0, 1, 2: m_AutoSave = Flag
		End Select
	End Property

	Private Sub Class_Initialize
		m_Error	   = -1
		m_Charset  = "utf-8"
		m_TotalSize= 0
		m_MaxSize  = 153600
		m_FileType = "jpg/gif/png"
		m_SavePath = ""
		m_AutoSave = 0
		Dim dtmNow : dtmNow = Date()
		m_strDate  = Year(dtmNow) & Right("0"&Month(dtmNow), 2) & Right("0" & Day(dtmNow), 2)
		m_lngTime  = Clng(Timer() * 1000)
		Set m_binForm = Server.CreateObject("ADODB.Stream")
		Set m_binItem = Server.CreateObject("ADODB.Stream")
		Set m_dicForm = Server.CreateObject("Scripting.Dictionary")
		m_dicForm.CompareMode = 1
	End Sub

	Private Sub Class_Terminate
		m_dicForm.RemoveAll
		Set m_dicForm = Nothing
		Set m_binItem = Nothing
		Set m_binForm = Nothing
	End Sub

	Public Function Open()
		Open = 0
		If m_Error = -1 Then m_Error = 0 Else Exit Function End If
		Dim lngRequestSize: lngRequestSize = Request.TotalBytes
		If m_TotalSize > 0 And lngRequestSize > m_TotalSize Then
			m_Error = 5: Exit Function
		ElseIf lngRequestSize < 1 Then
			m_Error = 4: Exit Function
		End If

		Dim lngChunkByte: lngChunkByte = 102400
		Dim lngReadSize: lngReadSize = 0
		m_binForm.Type = 1
		m_binForm.Open()
		Do
			m_binForm.Write(Request.BinaryRead(lngChunkByte))
			lngReadSize = lngReadSize + lngChunkByte
			If lngReadSize >= lngRequestSize Then Exit Do
		Loop		
		m_binForm.Position = 0
		Dim binRequestData: binRequestData = m_binForm.Read()

		Dim bCrLf, strSeparator, intSeparator
		bCrLf = ChrB(13) & ChrB(10)
		intSeparator = InstrB(1, binRequestData, bCrLf) - 1
		strSeparator = LeftB(binRequestData, intSeparator)

		Dim strItem, strInam, strFtyp, strPuri, strFnam, strFext, lngFsiz
		Const strSplit = "'"">"
		Dim strFormItem, strFileItem, intTemp, strTemp
		Dim p_start: p_start = intSeparator + 2
		Dim p_end
		Do
			p_end = InStrB(p_start, binRequestData, bCrLf & bCrLf) - 1
			m_binItem.Type = 1
			m_binItem.Open()
			m_binForm.Position = p_start
			m_binForm.CopyTo m_binItem, p_end - p_start
			m_binItem.Position = 0
			m_binItem.Type = 2
			m_binItem.Charset = m_Charset
			strItem = m_binItem.ReadText()
			m_binItem.Close()
			intTemp = Instr(39, strItem, """")
			strInam = Mid(strItem, 39, intTemp - 39)

			p_start = p_end + 4
			p_end = InStrB(p_start, binRequestData, strSeparator) - 1
			m_binItem.Type = 1
			m_binItem.Open()
			m_binForm.Position = p_start
			lngFsiz = p_end - p_start - 2
			m_binForm.CopyTo m_binItem, lngFsiz

			If Instr(intTemp, strItem, "filename=""") <> 0 Then
			If Not m_dicForm.Exists(strInam&"_From") Then
				strFileItem = strFileItem & strSplit & strInam
				If m_binItem.Size <> 0 Then
					intTemp = intTemp + 13
					strFtyp = Mid(strItem, Instr(intTemp, strItem, "Content-Type: ") + 14)
					strPuri = Mid(strItem, intTemp, Instr(intTemp, strItem, """") - intTemp)
					intTemp = InstrRev(strPuri, "\")
					strFnam = Mid(strPuri, intTemp + 1)
					m_dicForm.Add strInam&"_Type", strFtyp
					m_dicForm.Add strInam&"_Name", strFnam
					m_dicForm.Add strInam&"_Path", Left(strPuri, intTemp)
					m_dicForm.Add strInam&"_Size", lngFsiz
					If Instr(strFnam, ".") <> 0 Then strFext = Mid(strFnam, InstrRev(strFnam, ".") + 1) Else strFext = "" End If

					Select Case strFtyp
					Case "image/jpeg", "image/pjpeg", "image/jpg"
						If LCase(strFext) <> "jpg" Then strFext = "jpg"
						m_binItem.Position = 3
						Do While Not m_binItem.EOS
							Do
								intTemp = AscB(m_binItem.Read(1))
							Loop While intTemp = 255 And Not m_binItem.EOS
							
							If intTemp < 192 Or intTemp > 195 Then
								m_binItem.Read(Bin2Val(m_binItem.Read(2)) - 2)
							Else Exit Do End If
							
							Do
								intTemp = AscB(m_binItem.Read(1))
							Loop While intTemp < 255 And Not m_binItem.EOS
						Loop
						m_binItem.Read(3)
						m_dicForm.Add strInam&"_Height", Bin2Val(m_binItem.Read(2))
						m_dicForm.Add strInam&"_Width", Bin2Val(m_binItem.Read(2))
					Case "image/gif"
						If LCase(strFext) <> "gif" Then strFext = "gif"
						m_binItem.Position = 6
						m_dicForm.Add strInam&"_Width", BinVal2(m_binItem.Read(2))
						m_dicForm.Add strInam&"_Height", BinVal2(m_binItem.Read(2))
					Case "image/png"
						If LCase(strFext) <> "png" Then strFext = "png"
						m_binItem.Position = 18
						m_dicForm.Add strInam&"_Width", Bin2Val(m_binItem.Read(2))
						m_binItem.Read(2)
						m_dicForm.Add strInam&"_Height", Bin2Val(m_binItem.Read(2))
					Case "image/bmp"
						If LCase(strFext) <> "bmp" Then strFext = "bmp"
						m_binItem.Position = 18
						m_dicForm.Add strInam&"_Width", BinVal2(m_binItem.Read(4))
						m_dicForm.Add strInam&"_Height", BinVal2(m_binItem.Read(4))
					Case "application/x-shockwave-flash"
						If LCase(strFext) <> "swf" Then strFext = "swf"
						m_binItem.Position = 0
						If Ascb(m_binItem.Read(1)) = 70 Then
							m_binItem.Position = 8
							strTemp = Num2Str(Ascb(m_binItem.Read(1)), 2, 8)
							intTemp = Str2Num(Left(strTemp, 5), 2)
							strTemp = Mid(strTemp, 6)
							While (Len(strTemp) < intTemp * 4)
								strTemp = strTemp & Num2Str(Ascb(m_binItem.Read(1)), 2, 8)
							wend
							m_dicForm.Add strInam&"_Width", Int(Abs(Str2Num(Mid(strTemp, intTemp + 1, intTemp), 2) - Str2Num(Mid(strTemp, 1, intTemp), 2)) / 20)
							m_dicForm.Add strInam&"_Height", Int(Abs(Str2Num(Mid(strTemp, 3 * intTemp + 1, intTemp), 2) - Str2Num(Mid(strTemp, 2 * intTemp + 1, intTemp), 2)) / 20)
						End If
					End Select

					m_dicForm.Add strInam&"_Ext", strFext
					m_dicForm.Add strInam&"_From", p_start
					If m_AutoSave <> 2 Then
						intTemp = GetFerr(lngFsiz, strFext)
						m_dicForm.Add strInam&"_Err", intTemp
						If intTemp = 0 Then
							If m_AutoSave = 0 Then
								strFnam = GetTimeStr()
								If strFext <> "" Then strFnam = strFnam&"."&strFext
							End If
							m_binItem.SaveToFile Server.MapPath(m_SavePath & strFnam), 2
							m_dicForm.Add strInam, strFnam
						End If
					End If
				Else
					m_dicForm.Add strInam & "_Err", -1
				End If
			End If
			Else
				m_binItem.Position = 0
				m_binItem.Type = 2
				m_binItem.Charset = m_Charset
				strTemp = m_binItem.ReadText
				If m_dicForm.Exists(strInam) Then
					m_dicForm(strInam) = m_dicForm(strInam)&","&strTemp
				Else
					strFormItem = strFormItem & strSplit & strInam
					m_dicForm.Add strInam, strTemp
				End If
			End If

			m_binItem.Close()
			p_start = p_end + intSeparator + 2
		Loop Until p_start + 3 > lngRequestSize
		FormItem = Split(strFormItem, strSplit)
		FileItem = Split(strFileItem, strSplit)
		
		Open = lngRequestSize
	End Function

	Private Function GetTimeStr()
		m_lngTime = m_lngTime + 1
		GetTimeStr = m_strDate & Right("00000000"&m_lngTime, 8)
	End Function

	Private Function GetFerr(lngFsiz, strFext)
		Dim intFerr: intFerr = 0
		If lngFsiz > m_MaxSize And m_MaxSize > 0 Then
			If m_Error = 0 Or m_Error = 2 Then m_Error = m_Error + 1
			intFerr = intFerr+1
		End If
		If Instr(1, LCase("/"&m_FileType&"/"), LCase("/"&strFext&"/")) = 0 And m_FileType <> "" Then
			If m_Error < 2 Then m_Error = m_Error + 2
			intFerr = intFerr + 2
		End If
		GetFerr = intFerr
	End Function

	Public Function Save(Item, strFnam)
		Save = False
		If m_dicForm.Exists(Item&"_From") Then
			Dim intFerr, strFext
			strFext = m_dicForm(Item&"_Ext")
			intFerr = GetFerr(m_dicForm(Item&"_Size"), strFext)
			If m_dicForm.Exists(Item&"_Err") Then
				If intFerr = 0 Then m_dicForm(Item&"_Err") = 0 End If
			Else
				m_dicForm.Add Item&"_Err", intFerr
			End If
			If intFerr <> 0 Then Exit Function
			If VarType(strFnam) = 2 Then
				Select Case strFnam
					Case 0:strFnam = GetTimeStr()
						If strFext <> "" Then strFnam = strFnam&"."&strFext
					Case 1:strFnam = m_dicForm(Item&"_Name")
				End Select
			End If
			m_binItem.Type = 1
			m_binItem.Open
			m_binForm.Position = m_dicForm(Item&"_From")
			m_binForm.CopyTo m_binItem, m_dicForm(Item&"_Size")
			m_binItem.SaveToFile Server.MapPath(m_SavePath & strFnam), 2
			m_binItem.Close()
			If m_dicForm.Exists(Item) Then
				m_dicForm(Item) = strFnam
			Else
				m_dicForm.Add Item, strFnam
			End If
			Save = True
		End If
	End Function

	Public Function GetData(Item)
		GetData = ""
		If m_dicForm.Exists(Item&"_From") Then
			If GetFerr(m_dicForm(Item&"_Size"), m_dicForm(Item&"_Ext")) <> 0 Then Exit Function
			m_binForm.Position = m_dicForm(Item&"_From")
			GetData = m_binForm.Read(m_dicForm(Item&"_Size"))
		End If
	End Function

	Public Function Form(Item)
		If m_dicForm.Exists(Item) Then Form = m_dicForm(Item) Else Form = "" End If
	End Function

	Private Function BinVal2(bin)
		Dim lngValue: lngValue = 0
		Dim I: For I = LenB(bin) To 1 Step -1
			lngValue = lngValue *256 + AscB(MidB(bin, I, 1))
		Next
		BinVal2 = lngValue
	End Function

	Private Function Bin2Val(bin)
		Dim lngValue: lngValue = 0
		Dim I: For I = 1 To LenB(bin)
			lngValue = lngValue * 256 + AscB(MidB(bin, I, 1))
		Next
		Bin2Val = lngValue
	End Function

	Private Function Num2Str(num, base, lens)
		Dim I, ret: ret = ""
		While(num >= base)
			I = num Mod base
			ret = I & ret
			num = (num - I) / base
		wend
		Num2Str = Right(String(lens, "0") & num & ret, lens)
	End Function

	Private Function Str2Num(str, base)
		Dim ret: ret = 0 
		Dim I: For I = 1 To Len(str)
			ret = ret * base + Cint(Mid(str, I, 1))
		Next
		Str2Num = ret
	End Function
	
	Public Function Description(ByVal blError)
		Select Case blError
			Case -1: Description = "没有文件上传。"
			Case 0: Description = "上传成功。"
			Case 1: Description = "上传生效，文件大小超过了限制的 " & MaxSize / 1024 & "K，而未被保存。"
			Case 2: Description = "上传生效，文件类型受系统限制，而未被保存。"
			Case 3: Description = "上传生效，文件大小超过了限制的 " & MaxSize / 1024 & "K，且文件类型受系统限制，而未被保存。"
			Case 4: Description = "异常，不存在上传。"
			Case 5: Description = "异常，上传已经取消。"
		End Select
	End Function
	
End Class
%>
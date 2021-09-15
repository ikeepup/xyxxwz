<%
'=====================================================================
' ������ƣ�������վ����ϵͳ
' ��ǰ�汾��enchicms Version 3.0.0
' �������ڣ�2005-03-25
' �ٷ���վ���˳��ж�������Ƽ��������޹�˾(www.enchi.com.cn) 
' ����֧�֣����Ʒ�
' ���䣺liuyunfan@163.com
' QQ��21556923
'=====================================================================
' Copyright 2005-2008  All Rights Reserved.
'=====================================================================
Dim Myenchiasp,MyConn,IsConnection
IsConnection = False

Set Myenchiasp = New ClsProcess

Class ClsProcess
	Private CacheName, Reloadtime, LocalCacheName, Cache_Data
	Private MaxFileSize, sAllowExtName
	Public PathFileName, blnPassedTest
	Public PictureExist

	'-- ���ش�С����
	Public Property Let MaxSize(ByVal NewValue)
		MaxFileSize = NewValue * 1024
	End Property
	'-- ������������
	Public Property Let AllowExt(ByVal NewValue)
		sAllowExtName = NewValue
	End Property

	Public Property Get PictureEx()
		PictureEx = PictureExist
	End Property
	Public Property Get AllFileName()
		AllFileName = PathFileName
	End Property

	Private Sub Class_Initialize()
		On Error Resume Next
		Reloadtime = 28800
		CacheName = "myenchiasp"
		blnPassedTest = False
		PictureExist = False
		MaxFileSize = 0
		sAllowExtName = "gif|jpg|jpge|png|bmp|swf|fla|psd"
	End Sub

	Private Sub Class_Terminate()
		'-- Class_Terminate
	End Sub

	'===================���������沿�ֺ�����ʼ===================
	Public Property Let Name(ByVal vNewValue)
		LocalCacheName = LCase(vNewValue)
		Cache_Data = Application(CacheName & "_" & LocalCacheName)
	End Property
	Public Property Let Value(ByVal vNewValue)
		If LocalCacheName <> "" Then
			ReDim Cache_Data(2)
			Cache_Data(0) = vNewValue
			Cache_Data(1) = Now()
			Application.Lock
			Application(CacheName & "_" & LocalCacheName) = Cache_Data
			Application.UnLock
		Else
			Err.Raise vbObjectError + 1, "enchiaspCacheServer", " please change the CacheName."
		End If
	End Property
	Public Property Get Value()
		If LocalCacheName <> "" Then
			If IsArray(Cache_Data) Then
				Value = Cache_Data(0)
			Else
				'Err.Raise vbObjectError + 1, "enchiaspCacheServer", " The Cache_Data("&LocalCacheName&") Is Empty."
			End If
		Else
			Err.Raise vbObjectError + 1, "enchiaspCacheServer", " please change the CacheName."
		End If
	End Property
	Public Function ObjIsEmpty()
		ObjIsEmpty = True
		If Not IsArray(Cache_Data) Then Exit Function
		If Not IsDate(Cache_Data(1)) Then Exit Function
		If DateDiff("s", CDate(Cache_Data(1)), Now()) < (60 * Reloadtime) Then ObjIsEmpty = False
	End Function
	Public Sub DelCahe(MyCaheName)
		Application.Lock
		Application.Contents.Remove (CacheName & "_" & MyCaheName)
		Application.UnLock
	End Sub
	'===================���������沿�ֺ�������===================
	
	Public Function ChkBoolean(ByVal Values)
		If TypeName(Values) = "Boolean" Or IsNumeric(Values) Or LCase(Values) = "false" Or LCase(Values) = "true" Then
			ChkBoolean = CBool(Values)
		Else
			ChkBoolean = False
		End If
	End Function

	Public Function CheckNumeric(ByVal CHECK_ID)
		If CHECK_ID <> "" And IsNumeric(CHECK_ID) Then _
			CHECK_ID = CCur(CHECK_ID) _
		Else _
			CHECK_ID = 0
		CheckNumeric = CHECK_ID
	End Function

	Public Function ChkNumeric(ByVal CHECK_ID)
		If CHECK_ID <> "" And IsNumeric(CHECK_ID) Then
			CHECK_ID = CLng(CHECK_ID)
		Else
			CHECK_ID = 0
		End If
		ChkNumeric = CHECK_ID
	End Function

	Public Function CheckNull(ByVal str)
		If Not IsNull(str) And Trim(str) <> "" Then
			CheckNull = True
		Else
			CheckNull = False
		End If
	End Function

	Public Function CheckStr(ByVal str)
		If IsNull(str) Then
			CheckStr = ""
			Exit Function
		End If
		str = Replace(str, Chr(0), "")
		CheckStr = Replace(str, "'", "''")
	End Function

	Public Function CheckNostr(ByVal str)
		str = Trim(str)
		If Len(str) = 0 Then
			CheckNostr = ""
			Exit Function
		End If
		str = Replace(str, Chr(0), vbNullString)
		str = Replace(str, Chr(9), vbNullString)
		str = Replace(str, Chr(10), vbNullString)
		str = Replace(str, Chr(13), vbNullString)
		str = Replace(str, Chr(34), vbNullString)
		str = Replace(str, Chr(39), vbNullString)
		str = Replace(str, Chr(255), vbNullString)
		str = Replace(str, "%", "��")
		CheckNostr = Trim(str)
	End Function

	Public Function CheckNullStr(ByVal str)
		If Not IsNull(str) And Trim(str) <> "" And LCase(str) <> "http://" Then
			CheckNullStr = Trim(Replace(Replace(Replace(Replace(str, vbNewLine, ""), Chr(9), ""), Chr(39), ""), Chr(34), ""))
		Else
			CheckNullStr = ""
		End If
	End Function

	Public Function CheckMapPath(ByVal strPath)
		On Error Resume Next
		Dim fullPath
		strPath = Replace(Replace(Trim(strPath), "//", "/"), "\\", "\")

		If strPath = "" Then strPath = "."
		If InStr(strPath, ":") = 0 Then
			strPath = Replace(Trim(strPath), "\", "/")
			fullPath = Server.MapPath(strPath)
		Else
			strPath = Replace(Trim(strPath), "/", "\")
			fullPath = Trim(strPath)
		End If
		If Right(fullPath, 1) <> "\" Then fullPath = fullPath & "\"
		
		CheckMapPath = fullPath
	End Function
	Public Function ChkMapPath(ByVal strPath)
		On Error Resume Next
		Dim fullPath
		strPath = Replace(Replace(Trim(strPath), "//", "/"), "\\", "\")

		If strPath = "" Then strPath = "."
		If InStr(strPath, ":") = 0 Then
			strPath = Replace(Trim(strPath), "\", "/")
			fullPath = Server.MapPath(strPath)
		Else
			strPath = Replace(Trim(strPath), "/", "\")
			fullPath = Trim(strPath)
		End If
		If Right(fullPath, 1) <> "\" Then fullPath = fullPath & "\"
		fullPath = Left(fullPath, Len(fullPath) - 1)
		
		ChkMapPath = fullPath
	End Function
	'================================================
	'��������CheckRemoteUrl
	'��  �ã� �ж�Զ��URL
	'================================================
	Public Function CheckHTTP(ByVal URL)
		Dim Retrieval 
		
		On Error Resume Next
		Set Retrieval = CreateObject("MSXML2.XMLHTTP")
		With Retrieval
			.Open "HEAD", URL, False
			.send
			If .readyState <> 4 Then
				CheckHTTP = False
				Set Retrieval = Nothing
				Exit Function
			End If
			If .Status < 300 Then
				CheckHTTP = True
				Set Retrieval = Nothing
				Exit Function
			Else
				CheckHTTP = False
				Set Retrieval = Nothing
				Exit Function
			End If
		End With
		If Err.Number <> 0 Then
			CheckHTTP = False
			Err.Clear
			Set Retrieval = Nothing
			Exit Function
		End If
		Set Retrieval = Nothing
		Exit Function
	End Function
	'================================================
	'��������GetHTTPPage
	'��  �ã���ȡHTTPҳ
	'��  ����url   ----Զ��URL
	'����ֵ��Զ��HTML����
	'================================================
	Public Function GetRemoteData(ByVal URL, ByVal Cset)
		If Len(Cset) < 2 Then Cset = "GB2312"
		
		Dim strHeader
		Dim l
		
		On Error Resume Next
		
		Dim Retrieval
		Dim ObjStream
		Set ObjStream = CreateObject("ADODB.Stream")
		ObjStream.Type = 1
		ObjStream.Mode = 3
		ObjStream.Open
		Set Retrieval = CreateObject("MSXML2.XMLHTTP")
		With Retrieval
			.Open "GET", URL, False
			.setRequestHeader "Referer", URL
			.send
			If .readyState <> 4 Then Exit Function
			If .Status > 300 Then Exit Function
			'--��ȡĿ����վ�ļ�ͷ
			strHeader = .getResponseHeader("Content-Type")
			strHeader = UCase(strHeader)
			ObjStream.Write (.responseBody)
		End With
		Set Retrieval = Nothing
		
		If Len(strHeader) > 0 Then
			'--��ȡĿ���ļ�����
			l = InStrRev(strHeader, "CHARSET=", -1, 1)
			If l > 0 Then
				Cset = Right(strHeader, Len(strHeader) - l - 7)
			Else
				Cset = Cset
			End If
		End If

		ObjStream.Position = 0
		ObjStream.Type = 2
		ObjStream.Charset = Trim(Cset)
		GetRemoteData = ObjStream.ReadText
		ObjStream.Close
		Set ObjStream = Nothing
		Exit Function
	End Function
	'================================================
	'��������FindMatch
	'��  �ã���ȡ��ƥ�������
	'����ֵ����ȡ����ַ���
	'================================================
	Public Function FindMatch(ByVal str, ByVal start, ByVal last)
		
		Dim Match
		Dim s
		Dim FilterStr
		Dim MatchStr
		Dim strContent
		Dim ArrayFilter()
		Dim i, n
		Dim bRepeat
		
		If Len(start) = 0 Or Len(last) = 0 Then Exit Function
		
		On Error Resume Next
		
		MatchStr = "(" & CorrectPattern(start) & ")(.+?)(" & CorrectPattern(last) & ")"
		
		Dim re
		Set re = New RegExp
		re.IgnoreCase = True
		re.Global = True
		re.Pattern = MatchStr
		Set s = re.Execute(str)
		n = 0
		For Each Match In s
			If n = 0 Then
				n = n + 1
				ReDim ArrayFilter(n)
				ArrayFilter(n) = Match
			Else
				bRepeat = False
				For i = 0 To UBound(ArrayFilter)
					If UCase(Match) = UCase(ArrayFilter(i)) Then
						bRepeat = True
						Exit For
					End If
				Next
				If bRepeat = False Then
					n = n + 1
					ReDim Preserve ArrayFilter(n)
					ArrayFilter(n) = Match
				End If
			End If
		Next
		
		Set s = Nothing
		Set re = Nothing
		
		strContent = Join(ArrayFilter, "|||")
		strContent = Replace(strContent, start, "")
		strContent = Replace(strContent, last, "")
		
		FindMatch = Replace(strContent, "|||", vbNullString, 1, 1)
		Exit Function
	End Function
	'================================================
	'��������CutFixed
	'��  �ã���ȡ�̶����ַ���
	'��  ����strHTML   ----ԭ�ַ���
	'       start ------ ��ʼ�ַ���
	'       last ------ �����ַ���
	'================================================
	Public Function CutFixed(ByVal strHTML, ByVal start, ByVal last)
		Dim s
		Dim Match
		Dim strPattern
		Dim strContent
		Dim t, l

		t = Len(start): l = Len(last)
		If t = 0 Or l = 0 Then Exit Function

		strPattern = "(" & CorrectPattern(start) & ")(.+?)(" & CorrectPattern(last) & ")"

		On Error Resume Next

		Dim re
		Set re = New RegExp
		re.IgnoreCase = False
		re.Global = False
		re.Pattern = strPattern

		Set s = re.Execute(strHTML)
		For Each Match In s
			strContent = Match.Value
		Next

		Set s = Nothing
		Set re = Nothing
		CutFixed = Mid(strContent, t + 1, Len(strContent) - l - t)
		Exit Function
	End Function
	'================================================
	'��������CutFixate
	'����ֵ����ȡ����ַ���
	'================================================
	Public Function CutFixate(ByVal strHTML, ByVal start, ByVal last)
		
		Dim s
		Dim Match
		Dim strPattern
		Dim strContent
		Dim t, l

		t = Len(start): l = Len(last)
		If t = 0 Or l = 0 Then Exit Function

		strPattern = "(" & CorrectPattern(start) & ")(.+?)(" & CorrectPattern(last) & ")"

		On Error Resume Next

		Dim re
		Set re = New RegExp
		re.IgnoreCase = False
		re.Global = False
		re.Pattern = strPattern

		Set s = re.Execute(strHTML)
		For Each Match In s
			strContent = Match.Value
		Next

		Set s = Nothing
		Set re = Nothing
		
		CutFixate = Trim(strContent)
		Exit Function
	End Function
	'================================================
	'��������ReplaceTrim
	'��  �ã����˵��ַ������е�tab�ͻس��ͻ���
	'================================================
	Public Function ReplaceTrim(ByVal strContent)
		On Error Resume Next
		Dim re
		Set re = New RegExp
		re.IgnoreCase = True
		re.Global = True
		re.Pattern = "(" & Chr(8) & "|" & Chr(9) & "|" & Chr(10) & "|" & Chr(13) & ")"
		strContent = re.Replace(strContent, vbNullString)
		Set re = Nothing
		ReplaceTrim = strContent
		Exit Function
	End Function
	'================================================
	'��������ReplaceTrim
	'��  �ã����˵��ַ������е�tab�ͻس��ͻ���
	'================================================
	Public Function ReplacedTrim(ByVal strContent)
		On Error Resume Next
		Dim re
		Set re = New RegExp
		re.IgnoreCase = True
		re.Global = True
		re.Pattern = "(" & Chr(8) & "|" & Chr(9) & "|" & Chr(10) & "|" & Chr(13) & ")"
		strContent = re.Replace(strContent, vbNullString)
		re.Pattern = "(<!--(.+?)-->)"
		strContent = re.Replace(strContent, vbNullString)
		Set re = Nothing
		ReplacedTrim = strContent
		Exit Function
	End Function
	Public Function CheckMatch(ByVal strContent, ByVal start, ByVal last)
		If Len(strContent) = 0 Then Exit Function
		If Len(start) = 0 Then
			CheckMatch = strContent
			Exit Function
		End If
		If Len(last) = 0 Then
			CheckMatch = strContent
			Exit Function
		End If
		
		Dim strPattern
			
		On Error Resume Next
		
		strPattern = "(" & CorrectPattern(start) & ")(.+?)(" & CorrectPattern(last) & ")"
		
		Dim re
		Set re = New RegExp
		re.IgnoreCase = True
		re.Global = True
		re.Pattern = "(" & vbNewLine & ")"
		strContent = re.Replace(strContent, vbNullString)
		re.Pattern = strPattern
		strContent = re.Replace(strContent, vbNullString)
		Set re = Nothing
		CheckMatch = strContent
		Exit Function
	End Function
	Private Function CorrectPattern(ByVal str)
		str = Replace(str, "\", "\\")
		str = Replace(str, "~", "\~")
		str = Replace(str, "!", "\!")
		str = Replace(str, "@", "\@")
		str = Replace(str, "#", "\#")
		str = Replace(str, "%", "\%")
		str = Replace(str, "^", "\^")
		str = Replace(str, "&", "\&")
		str = Replace(str, "*", "\*")
		str = Replace(str, "(", "\(")
		str = Replace(str, ")", "\)")
		str = Replace(str, "-", "\-")
		str = Replace(str, "+", "\+")
		str = Replace(str, "[", "\[")
		str = Replace(str, "]", "\]")
		str = Replace(str, "<", "\<")
		str = Replace(str, ">", "\>")
		str = Replace(str, ".", "\.")
		str = Replace(str, "/", "\/")
		str = Replace(str, "?", "\?")
		str = Replace(str, "=", "\=")
		str = Replace(str, "|", "\|")
		str = Replace(str, "$", "\$")
		CorrectPattern = str
	End Function
	'================================================
	'��������ClearHtml
	'��  �ã����˵��ַ������е�HTML����
	'��  ����Str   ----ԭ�ַ���
	'����ֵ������ȡ����ַ���
	'================================================
	Public Function CheckHTML(ByVal str)
		On Error Resume Next
		
		Dim re
		Set re = New RegExp
		re.IgnoreCase = True
		re.Global = True
		re.Pattern = "<(.[^>]*)>"
		str = re.Replace(str, "")
		Set re = Nothing
		CheckHTML = str
		Exit Function

	End Function
	'================================================
	'��������Formatime
	'��  �ã���ʽ��ʱ��
	'================================================
	Public Function Formatime(ByVal datime)
		datime = Trim(Replace(Replace(Replace(Trim(datime), "&nbsp;", ""), Chr(255), ""), Chr(127), ""))
		datime = Trim(Replace(Replace(Replace(Replace(datime, vbNewLine, ""), Chr(9), ""), Chr(39), ""), Chr(34), ""))
		If Not IsDate(datime) Then
			Formatime = Now
			Exit Function
		End If
		If Len(datime) < 11 Then
			Formatime = CDate(datime & " " & FormatDateTime(Now, 3))
		Else
			Formatime = CDate(datime)
		End If
	End Function
	'================================================
	'��������GetRemoteUrl
	'��  �ã���ʽ����������URL
	'================================================
	Public Function FormatRemoteUrl(ByVal CurrentUrl, ByVal URL)
		Dim strUrl
		
		If Len(URL) < 2 Or Len(URL) > 255 Or Len(CurrentUrl) < 2 Then
			FormatRemoteUrl = vbNullString
			Exit Function
		End If

		CurrentUrl = Trim(Replace(Replace(Replace(Replace(Replace(CurrentUrl, "'", vbNullString), """", vbNullString), vbNewLine, vbNullString), "\", "/"), "|", vbNullString))
		URL = Trim(Replace(Replace(Replace(Replace(Replace(URL, "'", vbNullString), """", vbNullString), vbNewLine, vbNullString), "\", "/"), "|", vbNullString))
		
		If InStr(9, CurrentUrl, "/") = 0 Then
			strUrl = CurrentUrl
		Else
			strUrl = Left(CurrentUrl, InStr(9, CurrentUrl, "/") - 1)
		End If

		If strUrl = vbNullString Then strUrl = CurrentUrl
		Select Case Left(LCase(URL), 6)
			Case "http:/", "https:", "ftp://", "rtsp:/", "mms://"
				FormatRemoteUrl = URL
				Exit Function
		End Select

		If Left(URL, 1) = "/" Then
			FormatRemoteUrl = strUrl & URL
			Exit Function
		End If
		
		If Left(URL, 3) = "../" Then
			Dim ArrayUrl
			Dim ArrayCurrentUrl
			Dim ArrayTemp()
			Dim strTemp
			Dim i, n
			Dim c, l
			n = 0
			ArrayCurrentUrl = Split(CurrentUrl, "/")
			ArrayUrl = Split(URL, "../")
			c = UBound(ArrayCurrentUrl)
			l = UBound(ArrayUrl) + 1
			
			If c > l + 2 Then
				For i = 0 To c - l
					ReDim Preserve ArrayTemp(n)
					ArrayTemp(n) = ArrayCurrentUrl(i)
					n = n + 1
				Next
				strTemp = Join(ArrayTemp, "/")
			Else
				strTemp = strUrl
			End If
			URL = Replace(URL, "../", vbNullString)
			FormatRemoteUrl = strTemp & "/" & URL
			Exit Function
		End If
		
		strUrl = Left(CurrentUrl, InStrRev(CurrentUrl, "/"))
		FormatRemoteUrl = strUrl & Replace(URL, "./", vbNullString)
		Exit Function
	End Function
	'================================================
	'��������FormatContentUrl
	'��  �ã���ʽ��URL
	'��  ����Str   ----ԭ�ַ���
	'        url   ----��վURL
	'        ChildUrl   ----��Ŀ¼URL
	'����ֵ����ʽ��ȡ����ַ���
	'================================================
	Public Function FormatContentUrl(ByVal str, ByVal URL)
		Dim s_Content
		Dim re
		Dim ContentFile, ContentFileUrl
		Dim strTempUrl
		
		s_Content = str
		On Error Resume Next
		Set re = New RegExp
		re.IgnoreCase = True
		re.Global = True
		re.Pattern = "((src=|href=)((\S)+[.]{1}(" & sAllowExtName & ")))"
		Set ContentFile = re.Execute(s_Content)
		Dim sContentUrl(), n, i, bRepeat
		n = 0

		For Each ContentFileUrl In ContentFile
			If n = 0 Then
				n = n + 1
				ReDim sContentUrl(n)
				sContentUrl(n) = ContentFileUrl
			Else
				bRepeat = False
				For i = 1 To UBound(sContentUrl)
					If UCase(ContentFileUrl) = UCase(sContentUrl(i)) Then
						bRepeat = True
						Exit For
					End If
				Next
				If bRepeat = False Then
					n = n + 1
					ReDim Preserve sContentUrl(n)
					sContentUrl(n) = ContentFileUrl
				End If
			End If
		Next
		If n = 0 Then
			FormatContentUrl = s_Content
			Exit Function
		End If
		For i = 1 To n
			strTempUrl = Replace(Replace(Replace(Replace(sContentUrl(i), "src=", "", 1, -1, 1), "href=", "", 1, -1, 1), "'", ""), Chr(34), "")
			If LCase(Left(strTempUrl, 4)) <> "http" Then
				s_Content = Replace(s_Content, strTempUrl, FormatRemoteUrl(URL, strTempUrl), 1, -1, 1)
			End If
		Next
		Set re = Nothing
		PictureExist = True
		FormatContentUrl = s_Content
		Exit Function
	End Function
	'================================================
	'��������SaveRemoteFile
	'��  �ã�����Զ�̵��ļ�������
	'��  ����s_LocalFileName ------ �����ļ���
	'        s_RemoteFileUrl ------ Զ���ļ�URL
	'����ֵ��True  ----�ɹ�
	'        False ----ʧ��
	'================================================
	Public Function SaveRemoteFile(ByVal s_LocalFileName, ByVal s_RemoteFileUrl)

		Dim GetRemoteData
		Dim bError
		bError = False
		SaveRemoteFile = False
		On Error Resume Next
		
		Dim Retrieval
		Set Retrieval = CreateObject("MSXML2.XMLHTTP")
		
		With Retrieval
			.Open "GET", s_RemoteFileUrl, False, "", ""
			.setRequestHeader "Referer", s_RemoteFileUrl
			.send
			If .readyState <> 4 Then Exit Function
			If .Status > 300 Then Exit Function
			GetRemoteData = .responseBody
		End With
		Set Retrieval = Nothing
		
		If LenB(GetRemoteData) < 100 Then Exit Function
		If MaxFileSize > 0 Then
			If LenB(GetRemoteData) > MaxFileSize Then Exit Function
		End If
		
		Dim Ads
		Set Ads = Server.CreateObject("ADODB.Stream")
		With Ads
			.Type = 1
			.Open
			.Write GetRemoteData
			.SaveToFile ChkMapPath(s_LocalFileName), 2
			.Cancel
			.Close
		End With
		Set Ads = Nothing
		If Err.Number = 0 And bError = False Then
			SaveRemoteFile = True
		Else
			SaveRemoteFile = False
			Err.Clear
		End If
	End Function
	'================================================
	'��������RemoteToLocal
	'��  �ã��滻�ַ����е�Զ���ļ�Ϊ�����ļ�������Զ���ļ�
	'��  ����
	'       sHTML      : Ҫ�滻���ַ���
	'       sExt        : ִ���滻����չ��
	'================================================
	Public Function RemoteToLocal(ByVal sHTML, ByVal strPath)
		Dim s_Content
		Dim re
		Dim RemoteFile
		Dim RemoteFileUrl
		Dim SaveFileName
		Dim SaveFileType
		Dim a_RemoteUrl()
		Dim n
		Dim i
		Dim l
		Dim bRepeat
		Dim nFileNum
		Dim sContentPath
		s_Content = sHTML
		
		On Error Resume Next
		
		Set re = New RegExp
		re.IgnoreCase = True
		re.Global = True
		re.Pattern = "((http|https|ftp|rtsp|mms):(\/\/|\\\\){1}(([A-Za-z0-9_-])+[.]){1,}(net|com|cn|org|cc|tv|[0-9]{1,3})(\S*\/)((\S)+[.]{1}(" & sAllowExtName & ")))"
		Set RemoteFile = re.Execute(s_Content)
		n = 0
		'---- ת�����ظ�����
		For Each RemoteFileUrl In RemoteFile
			If n = 0 Then
				n = n + 1
				ReDim a_RemoteUrl(n)
				a_RemoteUrl(n) = RemoteFileUrl
			Else
				bRepeat = False
				For i = 1 To UBound(a_RemoteUrl)
					If UCase(RemoteFileUrl) = UCase(a_RemoteUrl(i)) Then
						bRepeat = True
						Exit For
					End If
				Next
				If bRepeat = False Then
					n = n + 1
					ReDim Preserve a_RemoteUrl(n)
					a_RemoteUrl(n) = RemoteFileUrl
				End If
			End If
		Next
		Set RemoteFile = Nothing
		Set re = Nothing
		If n = 0 Then
			PathFileName = ""
			RemoteToLocal = s_Content
			Exit Function
		End If
		'---- ��ʼ�滻����
		Dim UploadPath
		l = InStrRev(strPath, "UploadPic", -1)
		UploadPath = Right(strPath, Len(strPath) - l + 1)
		
		nFileNum = 0
		For i = 1 To n
			SaveFileType = Mid(a_RemoteUrl(i), InStrRev(a_RemoteUrl(i), ".") + 1)
			SaveFileName = GetRndFileName(SaveFileType)
			If SaveRemoteFile(strPath & SaveFileName, a_RemoteUrl(i)) = True Then
				nFileNum = nFileNum + 1
				If nFileNum > 0 Then
					PathFileName = PathFileName & "|"
				End If
				PathFileName = PathFileName & UploadPath & SaveFileName
				s_Content = Replace(s_Content, a_RemoteUrl(i), strPath & SaveFileName, 1, -1, 1)
			End If
		Next
		RemoteToLocal = s_Content
		Exit Function
	End Function
	Public Function FormatUrl(ByVal str)
		If Not IsNull(str) And Trim(str) <> "" And LCase(str) <> "http://" And Len(str) < 255 Then
			str = Trim(Replace(Replace(Replace(Replace(str, vbNewLine, ""), Chr(9), ""), Chr(39), ""), Chr(34), ""))
			If InStr(str, "://") > 0 Then
				FormatUrl = str
			Else
				FormatUrl = "http://" & str
			End If
		Else
			FormatUrl = ""
		End If
	End Function
	'--���ݹ���
	Public Function Html2Ubb(ByVal strContent, ByVal sRemoveCode)
		On Error Resume Next
		If Len(strContent) > 0 Then
			Dim ArrayCodes
			Dim re
			Set re = New RegExp
			If Len(sRemoveCode) < 21 Then sRemoveCode = "1|1|0|0|0|0|0|0|0|0|0|0"
			ArrayCodes = Split(sRemoveCode, "|")
			
			re.IgnoreCase = True
			re.Global = True
			
			'--���script�ű�
			If CInt(ArrayCodes(0)) = 1 Then
				re.Pattern = "(<s+cript(.+?)<\/s+cript>)"
				strContent = re.Replace(strContent, "")
			End If
			'--�������iframe���
			If CInt(ArrayCodes(1)) = 1 Then
				re.Pattern = "(<iframe(.+?)<\/iframe>)"
				strContent = re.Replace(strContent, "")
			End If
			'--�������object����
			If CInt(ArrayCodes(2)) = 1 Then
				re.Pattern = "(<object(.+?)<\/object>)"
				strContent = re.Replace(strContent, "")
			End If
			'--�������java applet
			If CInt(ArrayCodes(3)) = 1 Then
				re.Pattern = "(<applet(.+?)<\/applet>)"
				strContent = re.Replace(strContent, "")
			End If
			'--�������div��ǩ
			If CInt(ArrayCodes(4)) = 1 Then
				re.Pattern = "(<DIV(.+?)>)"
				strContent = re.Replace(strContent, "")
				re.Pattern = "(<\/DIV>)"
				strContent = re.Replace(strContent, "")
			End If
			'--�������font��ǩ
			If CInt(ArrayCodes(5)) = 1 Then
				re.Pattern = "(<FONT(.+?)>)"
				strContent = re.Replace(strContent, "")
				re.Pattern = "(<\/FONT>)"
				strContent = re.Replace(strContent, "")
			End If
			'--�������span��ǩ
			If CInt(ArrayCodes(6)) = 1 Then
				re.Pattern = "(<SPAN(.+?)>)"
				strContent = re.Replace(strContent, "")
				re.Pattern = "(<\/SPAN>)"
				strContent = re.Replace(strContent, "")
			End If
			'--�������A��ǩ
			If CInt(ArrayCodes(7)) = 1 Then
				re.Pattern = "(<A(.+?)>)"
				strContent = re.Replace(strContent, "")
				re.Pattern = "(<\/A>)"
				strContent = re.Replace(strContent, "")
			End If
			'--�������img��ǩ
			If CInt(ArrayCodes(8)) = 1 Then
				re.Pattern = "(<IMG(.+?)>)"
				strContent = re.Replace(strContent, "")
			End If
			'--�������FORM��ǩ
			If CInt(ArrayCodes(9)) = 1 Then
				re.Pattern = "(<FORM(.+?)>)"
				strContent = re.Replace(strContent, "")
				re.Pattern = "(<\/FORM>)"
				strContent = re.Replace(strContent, "")
			End If
			'--�������HTML��ǩ
			If CInt(ArrayCodes(10)) = 1 Then
				re.Pattern = "<(.[^>]*)>"
				strContent = re.Replace(strContent, "")
			End If
			re.Pattern = "(" & Chr(8) & "|" & Chr(9) & "|" & Chr(10) & "|" & Chr(13) & ")"
			strContent = re.Replace(strContent, vbNullString)
			re.Pattern = "(<!--(.+?)-->)"
			strContent = re.Replace(strContent, vbNullString)
			re.Pattern = "(<TBODY>)"
			strContent = re.Replace(strContent, "")
			re.Pattern = "(<\/TBODY>)"
			strContent = re.Replace(strContent, "")
			re.Pattern = "(<" & Chr(37) & ")"
			strContent = re.Replace(strContent, "&lt;%")
			re.Pattern = "(" & Chr(37) & ">)"
			strContent = re.Replace(strContent, "%&gt;")
			Set re = Nothing
			Html2Ubb = strContent
		Else
			Html2Ubb = ""
		End If
		Exit Function
	End Function
	'--���������滻
	Public Function ReplaceClass(ByVal ClassName, ByVal ClassList)
		If Len(ClassList) < 3 Then
			ReplaceClass = Trim(ClassName)
			Exit Function
		End If
		ClassName = Trim(ClassName)
		If Len(ClassName) = 0 Then Exit Function
		
		Dim i
		Dim ArrayClassList
		Dim ArrayClassName
		
		On Error Resume Next
		
		ArrayClassList = Split(ClassList, "$$$")
		For i = 0 To UBound(ArrayClassList)
			If Len(ArrayClassList(i)) > 2 Then
				ArrayClassName = Split(ArrayClassList(i), "|")
				ClassName = Replace(ClassName, ArrayClassName(0), ArrayClassName(1))
			End If
		Next
		ReplaceClass = ClassName
	End Function
	'��ʽ���ļ���СKB
	Public Function FormatSize(ByVal strFileSize)
		On Error Resume Next
		Dim valFileSize
		strFileSize = Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(UCase(strFileSize), "��", "K"), "��", "B"), "��", "M"), "��", "G"), "��", "Y"), "��", "T"), "��", "E"), "��", "S")
		valFileSize = Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(UCase(strFileSize), "BYTE", ""), "K", ""), "M", ""), "G", ""), "B", ""), "S", ""), " ", ""), "&NBSP;", ""), vbNewLine, ""), Chr(-24159), ""), Chr(9), ""), Chr(11), "")
		If IsNumeric(valFileSize) Then
			If InStr(strFileSize, "K") > 0 Then
				valFileSize = valFileSize
			ElseIf InStr(strFileSize, "M") > 0 Then
				valFileSize = valFileSize * 1024
			ElseIf InStr(strFileSize, "G") > 0 Then
				valFileSize = valFileSize * 1024 * 1024
			ElseIf InStr(strFileSize, "BYTE") > 0 Then
				valFileSize = valFileSize \ 1024
			Else
				valFileSize = valFileSize
			End If
		Else
			valFileSize = 0
		End If
		FormatSize = valFileSize
		Exit Function
	End Function
	'--��������Ŀ¼
	Public Function BuildDatePath(ByVal DirForm)
		On Error Resume Next
		DirForm = CInt(DirForm)
		Dim DatePath
		Select Case DirForm
		Case 1
			DatePath = Year(Now) & "-" & Month(Now)
			BuildDatePath = DatePath & "/"
		Case 2
			DatePath = Year(Now) & "_" & Month(Now)
			BuildDatePath = DatePath & "/"
		Case 3
			DatePath = Year(Now) & Month(Now)
			BuildDatePath = DatePath & "/"
		Case 4
			DatePath = Year(Now)
			BuildDatePath = DatePath & "/"
		Case 5
			DatePath = Year(Now) & "/" & Month(Now)
			BuildDatePath = DatePath & "/"
		Case 6
			DatePath = Year(Now) & "/" & Month(Now) & "/" & Day(Now)
			BuildDatePath = DatePath & "/"
		Case 7
			DatePath = Year(Now) & Month(Now) & Day(Now)
			BuildDatePath = DatePath & "/"
		Case Else
			BuildDatePath = vbNullString
		End Select
	End Function
	'================================================
	'��������GetRndFileName
	'��  �ã�ȡ����ļ���
	'��  ����sExt   ----ԭ�ַ���
	'����ֵ����ȡ����ļ���
	'================================================
	Public Function GetRndFileName(ByVal sExt)
		Dim sRnd
		Randomize
		sRnd = Int(900 * Rnd) + 100
		GetRndFileName = Year(Now) & Month(Now) & Day(Now) & Hour(Now) & Minute(Now) & Second(Now) & sRnd & "." & sExt
	End Function
	'=================================================
	'��������GetFileExtName
	'��  �ã���ȡ�ļ���չ��
	'=================================================
	Public Function GetFileExtName(ByVal sName)
		Dim FileName
		FileName = Split(sName, ".")
		GetFileExtName = FileName(UBound(FileName))
	End Function
	'================================================
	'��������GetRndHits
	'��  �ã�ȡ��������
	'================================================
	Public Function GetRndHits()
		Dim sRnd
		Randomize
		sRnd = Int(900 * Rnd) + 100
		GetRndHits = sRnd
	End Function
	Public Function CheckPath(ByVal sPath)
		'-- �����ļ�·��
		sPath = Trim(sPath)
		If Right(sPath, 1) <> "\" And sPath <> "" Then
			sPath = sPath & "\"
		End If
		CheckPath = sPath
	End Function
	'================================================
	'��������CreatedPathEx
	'��  �ã�FSO�����༶Ŀ¼
	'��  ����LocalPath   ----ԭ�ļ�·��
	'����ֵ��False  ----  True
	'================================================
	Public Function CreatedPathEx(ByVal sPath)
		sPath = Replace(sPath, "/", "\")
		sPath = Replace(sPath, "\\", "\")
		On Error Resume Next
		
		Dim strHostPath,strPath
		Dim sPathItem,sTempPath
		Dim i,fso
		
		Set fso = Server.CreateObject("Scripting.FileSystemObject")
		strHostPath = Server.MapPath("/")
		If InStr(sPath, ":") = 0 Then sPath = Server.MapPath(sPath)
		If fso.FolderExists(sPath) Or Len(sPath) < 3 Then
			CreatedPathEx = True
			Exit Function
		End If
		
		strPath = Replace(sPath, strHostPath, vbNullString,1,-1,1)
		sPathItem = Split(strPath, "\")
		
		If InStr(LCase(sPath), LCase(strHostPath)) = 0 Then
			sTempPath = sPathItem(0)
		Else
			sTempPath = strHostPath
		End If
		
		For i = 1 To UBound(sPathItem)
			If sPathItem(i) <> "" Then
				sTempPath = sTempPath & "\" & sPathItem(i)
				If fso.FolderExists(sTempPath) = False Then
					fso.CreateFolder sTempPath
				End If
			End If
		Next
		Set fso = Nothing
		If Err.Number <> 0 Then Err.Clear
		CreatedPathEx = True
	End Function
	
	'--ɾ���ļ�
	Public Function DeleteFiles(ByVal sFilePath)
		On Error Resume Next
		Dim fso
		Set fso = Server.CreateObject("Scripting.FileSystemObject")
		fso.DeleteFile sFilePath, True
		DeleteFiles = True
		Set fso = Nothing
		Exit Function
	End Function
	'=============================================================
	'��������ChkFormStr
	'��  �ã����˱��ַ�
	'��  ����str   ----ԭ�ַ���
	'����ֵ�����˺���ַ���
	'=============================================================
	Public Function FormatStr(ByVal str)
		Dim fString
		fString = str
		If Len(str) = 0 Then
			FormatStr = ""
			Exit Function
		End If
		fString = Replace(fString, "'", "&#39;")
		fString = Replace(fString, Chr(34), "&quot;")
		fString = Replace(fString, Chr(13), "")
		fString = Replace(fString, Chr(10), "")
		fString = Replace(fString, Chr(9), "")
		fString = Replace(fString, ">", "&gt;")
		fString = Replace(fString, "<", "&lt;")
		fString = Replace(fString, "%", "��")
		FormatStr = Trim(fString)
	End Function

End Class

Public Sub OutErrors(msg)
	Response.Write "<script language=""javascript"">" & vbCrLf
	Response.Write "alert(""" & Replace(Replace(Replace(msg, "<li>", "", 1, -1, 1), "</li>", "\n", 1, -1, 1), """", "\""") & """);"
	Response.Write "history.back();" & vbCrLf
	Response.Write "</script>" & vbCrLf
	Response.Flush
End Sub
Public Sub OutScript(msg)
	Response.Write "<script language=""javascript"">" & vbCrLf
	Response.Write "alert(""" & Replace(Replace(Replace(msg, "<li>", "", 1, -1, 1), "</li>", "\n", 1, -1, 1), """", "\""") & """);"
	Response.Write "location.replace(""" & Request.ServerVariables("HTTP_REFERER") & """);" & vbCrLf
	Response.Write "</script>" & vbCrLf
	Response.Flush: Response.End
End Sub
Public Sub ReturnError(ErrMsg)
	Response.Write "<br><br><table cellpadding=5 cellspacing=1 border=0 align=center class=tableBorder1>" & vbCrLf
	Response.Write "  <tr><th colspan=2>������ʾ��Ϣ!</th></tr>" & vbCrLf
	Response.Write "  <tr><td colspan=2 align=center height=50 class=TableRow1>" & ErrMsg & "</td></tr>" & vbCrLf
	Response.Write "</table><br>" & vbCrLf
	Response.Flush
End Sub
'================================================
'��������ShowListPage
'��  �ã�ͨ�÷�ҳ
'================================================
Public Function ShowListPage(ByVal CurrentPage, ByVal Pcount, ByVal totalrec, ByVal PageNum, ByVal strLink, ByVal ListName)
	With Response
		.Write "<script>"
		.Write "ShowListPage("
		.Write CurrentPage
		.Write ","
		.Write Pcount
		.Write ","
		.Write totalrec
		.Write ","
		.Write PageNum
		.Write ",'"
		.Write strLink
		.Write "','"
		.Write ListName
		.Write "');"
		.Write "</script>" & vbNewLine
	End With
End Function
'-- �������ݿ�
Sub DatabaseConnection()
	On Error Resume Next
	Set MyConn = Server.CreateObject("ADODB.Connection")
	MyConn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & ChkMapPath(DBPath)
	If Err Then
		Err.Clear
		Set MyConn = Nothing
		Response.Write "���ݿ����ӳ������conn.asp���ɼ����ݿ������ִ���"
		Response.End
	End If
	IsConnection = True
End Sub
%>
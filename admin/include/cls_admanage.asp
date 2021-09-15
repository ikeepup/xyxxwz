<%
'=====================================================================
' 软件名称：恩池网站管理系统
' 当前版本：enchicms Version 3.0.0
' 更新日期：2005-03-25
' 官方网站：运城市恩池软件科技开发有限公司(www.enchi.com.cn) 
' 技术支持：柳云帆
' 邮箱：liuyunfan@163.com
' QQ：21556923
'=====================================================================
' Copyright 2005-2008  All Rights Reserved.
'=====================================================================
Class Admanage_Cls
	Private boardid
	Private JsFileName
	Private maxadnum

	Private Sub Class_Initialize()
		On Error Resume Next
		boardid = 1
	End Sub
	
	Public Property Let adboardid(ByVal NewValue)
		boardid = CLng(NewValue)
	End Property
	Private Sub LoadAdBoardInfo()
		Dim Rs
		On Error Resume Next
		Set Rs = enchiasp.Execute("SELECT fileName,maxads FROM ECCMS_Adboard WHERE boardid=" & boardid)
		If Rs.BOF And Rs.EOF Then
			JsFileName = "../adfile/ad.js"
			maxadnum = 1
		Else
			JsFileName = "../adfile/" & Trim(Rs("fileName"))
			maxadnum = Rs("maxads")
		End If
		Set Rs = Nothing
	End Sub

	Private Function ReadFlashAndPic(ByVal url, ByVal Picurl, _
		ByVal width, ByVal height, _
		ByVal Readme, ByVal isFlash)
		
		Dim strTemp
		If CInt(isFlash) = 1 Then
			
			
		strTemp = "<object classid='clsid:D27CDB6E-AE6D-11CF-96B8-444553540000' id='obj1' codebase='http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,40,0' border='0' width='" & width & "' height='" & height & "'>"
		strTemp =strTemp& "<param name='movie' value='" & Picurl & "'>"
		strTemp =strTemp& "	<param name='quality' value='High'>"
		strTemp =strTemp& "	<param name='wmode' value='transparent'>"
		strTemp =strTemp& "	<embed src='" & Picurl & "' pluginspage='http://www.macromedia.com/go/getflashplayer' type='application/x-shockwave-flash' name='obj1' width='" & width & "' height='" & height & "' wmode='transparent'></object>"

			
			'strTemp = "<embed src='" & Picurl & "' quality='high' pluginspage='http://www.macromedia.com/go/getflashplayer' type='application/x-shockwave-flash' width='" & width & "' height='" & height & "'></embed>"
		
		Else
			strTemp = "<a href='" & url & "' target='_blank'><img src='" & Picurl & "' width='" & width & "' height='" & height & "' border=0 alt='" & fixjs(Readme) & "'></a>"
		End If
		ReadFlashAndPic = strTemp
	End Function

	Public Function fixjs(ByVal str)
		If str <> "" Then
			str = Replace(str, "\", "\\")
			str = Replace(str, Chr(34), "\""")
			str = Replace(str, Chr(39), "\'")
			str = Replace(str, Chr(13), "")
			str = Replace(str, Chr(10), "")
			str = Replace(str, vbNewLine, vbNullString)
		End If
		fixjs = str
		Exit Function
	End Function

	Public Sub CreateJsFile()
		Dim Rs, SQL, strTemp, i
		Dim strFalshAndPic, strAdContent, strMargin
		Dim strCommon
		Dim strFloat
		Dim strFixed2
		Dim strFixed3
		Dim strRunCode
		Dim strAdcode
		
		strMargin = ""
		Call LoadAdBoardInfo
		On Error Resume Next
		Set Rs = enchiasp.Execute("SELECT TOP " & maxadnum & " * FROM ECCMS_Adlist WHERE isLock=0 And boardid=" & boardid & " ORDER BY startime DESC")
		If Not (Rs.BOF And Rs.EOF) Then
			i = 0
			Do While Not Rs.EOF
				i = i + 1
				
				Select Case CInt(Rs("flag"))
				Case 1
					strFalshAndPic = ReadFlashAndPic(Rs("url"), enchiasp.ReadFileUrl(Rs("Picurl")), Rs("width"), Rs("height"), Rs("Readme"), Rs("isFlash"))
					strFloat = strFalshAndPic
				Case 2, 3
					If Rs("flag") = 3 Then strMargin = "style='right:" & Rs("sidemargin") & "px;POSITION:absolute;TOP:" & Rs("topmargin") & "px;'"
					If Rs("flag") = 2 Then strMargin = "style='left:" & Rs("sidemargin") & "px;POSITION:absolute;TOP:" & Rs("topmargin") & "px;'"
					strFalshAndPic = ReadFlashAndPic(Rs("url"), enchiasp.ReadFileUrl(Rs("Picurl")), Rs("width"), Rs("height"), Rs("Readme"), Rs("isFlash"))
					strFixed2 = strFixed2 & "document.all.lovexin" & Rs("id") & ".style.pixelTop+=percent;" & vbNewLine
					strFixed3 = strFixed3 & "suspendcode" & Rs("id") & "=""<div id=lovexin" & Rs("id") & " " & strMargin & ">" & strFalshAndPic & "</div>""" & vbNewLine & "document.write(suspendcode" & Rs("id") & "); " & vbNewLine
				Case 4
					'strRunCode = strRunCode & vbNewLine & "window.open(""" & enchiasp.SiteUrl & enchiasp.InstallDir & "runads.asp?id=" & Rs("id") & """,""runads" & Rs("id") & """,""toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no,width=" & Rs("Width") & ",height=" & Rs("Height") & ",top=" & Rs("topmargin") & ",left=" & Rs("sidemargin") & """);" & vbNewLine
					strRunCode = strRunCode & vbNewLine & "window.open(""" &  Rs("url") & """," & Rs("Readme") & ",""toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no,width=" & Rs("Width") & ",height=" & Rs("Height") & ",top=" & Rs("topmargin") & ",left=" & Rs("sidemargin") & """);" & vbNewLine

				Case 5
					strAdContent = strAdContent & vbNewLine & Rs("Adcode") & vbNewLine
					strAdcode = vbNewLine & "document.writeln(""<iframe scrolling='no' frameborder='0' marginheight='0' marginwidth='0' width='" & Rs("width") & "' height='" & Rs("height") & "' src=" & enchiasp.SiteUrl & enchiasp.InstallDir & "adfile/ad" & boardid & ".htm></iframe>"");" & vbNewLine
				Case Else
					strFalshAndPic = ReadFlashAndPic(Rs("url"), enchiasp.ReadFileUrl(Rs("Picurl")), Rs("width"), Rs("height"), Rs("Readme"), Rs("isFlash"))
					strCommon = strCommon & "document.writeln(""" & strFalshAndPic & """);" & vbNewLine
				End Select
				Rs.MoveNext
			Loop
		End If
		Rs.Close: Set Rs = Nothing
		strTemp = strCommon
		If Trim(strFloat) <> "" Then
			strTemp = strTemp & enchiasp.Readfile("include/float.inc")
			strTemp = Replace(strTemp, "{$FloatCode}", strFloat)
		End If
		If Trim(strFixed2) <> "" Then
			strTemp = strTemp & enchiasp.Readfile("include/fixed.inc")
			strTemp = Replace(strTemp, "{$Scroll}", strFixed2)
			strTemp = Replace(strTemp, "{$SuspendCode}", strFixed3)
		End If
		If Trim(strAdcode) <> "" Then
			Dim strHtml, HtmlName
			HtmlName = "../adfile/ad" & boardid & ".htm"
			strHtml = enchiasp.Readfile("include/html.inc")
			strHtml = Replace(strHtml, "{$HtmlContent}", strAdContent)
			enchiasp.CreatedTextFile HtmlName, strHtml
			strTemp = strTemp & strAdcode
		End If
		strTemp = strTemp & strRunCode
		enchiasp.CreatedTextFile JsFileName, strTemp
	End Sub
End Class
%>
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
Dim enchicms
Set enchicms = New Cls_AdminUploadFile

Class Cls_AdminUploadFile
	Private fromPath, modules
	Private ChannelDir, fullPath, FilePath, UploadDir, ThisDir
	Private Action, AdminFlag,rsChannel



	Public Sub ShowUploadFile()
		if Request("ChannelID")="-1" then
			ChannelID=-1
		else
		ChannelID = enchiasp.ChkNumeric(Request("ChannelID"))
		end if
		AdminFlag = "AdminUpload" & ChannelID
		Action = LCase(Request("action"))
		
		If Not ChkAdmin(AdminFlag) Then
		'response.write adminflag
		'response.end
			Server.Transfer ("showerr.asp")
			Response.End
		End If
		Admin_header
		If ChannelID > 0 Then
			Set rsChannel = enchiasp.Execute("SELECT ChannelDir,modules FROM ECCMS_Channel WHERE ChannelType < 2 And ChannelID = " & ChannelID)
			If Not (rsChannel.BOF And rsChannel.EOF) Then
				ChannelDir = Trim(enchiasp.InstallDir) & Trim(rsChannel("ChannelDir"))
				modules = rsChannel("modules")
			Else
				ChannelDir = Trim(enchiasp.InstallDir) & "adfile/"
				modules = 0
			End If
			rsChannel.Close: Set rsChannel = Nothing
		Else
			if ChannelID=-1 then
				ChannelDir = Trim(enchiasp.InstallDir) & "fengmian/"
				modules = 0

			else
			ChannelID = 0
			modules = 0
			ChannelDir = Trim(enchiasp.InstallDir) & "adfile/"
			end if
		End If
		If Trim(Request("UploadDir")) <> "" Then
			UploadDir = Trim(Request("UploadDir")) & "/"
		End If
		If Trim(Request("ThisDir")) <> "" Then
			ThisDir = Trim(Request("ThisDir")) & "/"
		End If
		ThisDir = Replace(ThisDir, "\", "/")
		If ChannelID = 0 Then
			fromPath = Replace("adfile/" & UploadDir, "\", "/")
		Else
			fromPath = Replace(UploadDir, "\", "/")
		End If
		FilePath = Replace(ChannelDir & UploadDir, "\", "/")
		fullPath = Server.MapPath(FilePath)

		Select Case Trim(Action)
		Case "clear"
			Call ClearUploadFile
		Case "delete"
			Call DelUselessFile
		Case "del"
			Call DelFile
		Case "delalldirfile"
			Call DelAllDirFile
		Case "delthisallfile"
			Call DelThisAllFile
		Case "delemptyfolder"
			Call DelEmptyFolder
		Case Else
			Call ShowUploadMain
		End Select
		If FoundErr = True Then
			ReturnError (ErrMsg)
		End If
		Admin_footer
	End Sub
	'=================================================
	'��������ShowSelectFile
	'��  �ã���ʾѡ���ļ�
	'=================================================
	Public Sub ShowSelectFile()
		Admin_header
		Response.Write "<base target=""_self"">" & vbNewLine
		if Request("ChannelID")="-1" then
			ChannelID=CInt(Request("ChannelID"))
		else
			ChannelID = enchiasp.ChkNumeric(Request("ChannelID"))
		end if
		AdminFlag = "AdminSelect" & ChannelID
		If Not ChkAdmin(AdminFlag) Then
			Server.Transfer ("showerr.asp")
			Response.End
		End If
		If ChannelID > 0 Then
			ChannelID = CInt(Request("ChannelID"))
			Set rsChannel = enchiasp.Execute("SELECT ChannelDir FROM ECCMS_Channel WHERE ChannelType < 2 And ChannelID = " & ChannelID)
			If Not (rsChannel.BOF And rsChannel.EOF) Then
				ChannelDir = Trim(enchiasp.InstallDir) & Trim(rsChannel("ChannelDir"))
			Else
				ChannelDir = Trim(enchiasp.InstallDir) & "adfile/"
			End If
			rsChannel.Close: Set rsChannel = Nothing
		Elseif  ChannelID = -1 then
				
				ChannelID = -1
				ChannelDir = Trim(enchiasp.InstallDir) & "fengmian/"
		else
				ChannelID = 0
				ChannelDir = Trim(enchiasp.InstallDir) & "adfile/"
		End If
		
		
		If Trim(Request("UploadDir")) <> "" Then
			UploadDir = Trim(Request("UploadDir")) & "/"
		End If
		'If Trim(Request("ThisDir")) <> "" Then
			'ThisDir = Trim(Request("ThisDir")) & "/"
		'End If
		'ThisDir = Replace(ThisDir, "\", "/")
		If ChannelID = 0 Then
			fromPath = Replace("adfile/" & UploadDir, "\", "/")
		Else
			fromPath = Replace(UploadDir, "\", "/")
		End If
		FilePath = Replace(ChannelDir & UploadDir, "\", "/")
		fullPath = Server.MapPath(FilePath)
		Call ShowSelectMain
		If FoundErr = True Then
			ReturnError (ErrMsg)
		End If
		Admin_footer
	End Sub
	'=================================================
	'��������ShowSelectMain
	'��  �ã���ʾѡ���ļ���ҳ��
	'=================================================
	Private Sub ShowSelectMain()
		Dim maxperpage, CurrentPage, TotalNumber, Pcount
		Dim fso, FileCount, TotleSize, totalPut
		maxperpage = 20 '###ÿҳ��ʾ��
		
		If IsNumeric(Request("page")) And Trim(Request("page")) <> "" Then
			CurrentPage = CLng(Request("page"))
		Else
			CurrentPage = 1
		End If
		If CLng(CurrentPage) = 0 Then CurrentPage = 1
		On Error Resume Next
		If Not IsObjInstalled(enchiasp.FSO_ScriptName) Then
			Response.Write "<b><font color=red>��ķ�������֧�� fso(Scripting.FileSystemObject)! ����ʹ�ñ�����</font></b>"
		End If
		
		Response.Write "<table border=0 align=center cellpadding=3 cellspacing=1 class=tableborder>"
		Response.Write "<tr>"
		Response.Write "        <th colspan=""2"">�ļ�Ŀ¼</th>"
		Response.Write "</tr>"
		Response.Write "<tr>"
		Response.Write "        <td class=tablerow1 colspan=""2"">"
		Call ShowChildFolder
		Response.Write "</td>"
		Response.Write "</tr>"
		Response.Write "<tr>"
		Response.Write "        <td width=""50%"" class=tablerow2>��ǰĿ¼��" & FilePath & "</td>"
		Response.Write "        <td width=""50%"" align=center class=tablerow2>"
		If Trim(Request("ThisDir")) <> "" Then
			Response.Write "<a href=""admin_selFile.asp?ChannelID=" & ChannelID & "&UploadDir=" & Left(Request("UploadDir"),Len(Request("UploadDir"))-Len(Mid(Request("UploadDir"), InStrRev(Request("UploadDir"), "/")))) & "&ThisDir=" & Request("ThisDir") & """>��������һ��Ŀ¼</a>"
		End If
		Response.Write "</td>"
		Response.Write "</tr>"
		Response.Write "</table><br>" & vbNewLine

		Set fso = CreateObject(enchiasp.FSO_ScriptName)
		If fso.FolderExists(fullPath) Then
			Dim fsoFile, fsoFileSize
			Dim DirFiles, DirFolder
			Set fsoFile = fso.GetFolder(fullPath)
			'fsoFileSize = fsoFile.size '�ռ��Сͳ��
			Dim c
			FileCount = fsoFile.Files.Count
			TotleSize = GetFileSize(fsoFile.Size)
			totalPut = fsoFile.Files.Count
			If CurrentPage < 1 Then
				CurrentPage = 1
			End If
			If (CurrentPage - 1) * maxperpage > totalPut Then
				If (totalPut Mod maxperpage) = 0 Then
					CurrentPage = totalPut \ maxperpage
				Else
					CurrentPage = totalPut \ maxperpage + 1
				End If
			End If
			FileCount = 0
			c = 0
			Response.Write "<table border=0 align=center cellpadding=3 cellspacing=1 class=tableborder>" & vbNewLine
			Response.Write "<tr><td colspan=4 class=tablerow1>" & vbNewLine
			Response.Write showpage(CurrentPage, totalPut, maxperpage, TotleSize)
			Response.Write "</td></tr>" & vbNewLine
			Response.Write "<tr>" & vbNewLine
			For Each DirFiles In fsoFile.Files
				c = c + 1
				If c > maxperpage * (CurrentPage - 1) Then
					Response.Write "<td class=tablerow2>"

						'Response.Write "<div align=center><a href='#' onClick=""window.returnValue='" &  GetFilePic(FilePath & DirFiles.Name) & "|" & CLng(DirFiles.Size \ 1024) & "';window.close();""><img src='" & GetFilePic(FilePath & DirFiles.Name) & "' width=140 height=100 border=0 alt='���ͼƬ�����أ���������ļ������鿴ԭʼ�ļ���'></a></div>"


						Response.Write "<div align=center><a href='#' onClick=""window.returnValue='" & fromPath & DirFiles.Name & "|" & CLng(DirFiles.Size \ 1024) & "';window.close();""><img src='" & GetFilePic(FilePath & DirFiles.Name) & "' width=140 height=100 border=0 alt='���ͼƬ�����أ���������ļ������鿴ԭʼ�ļ���'></a></div>"

					Response.Write "�ļ�����<a href='" & FilePath & DirFiles.Name & "'target=_blank>" & DirFiles.Name & "</a><br>"
					Response.Write "�ļ���С��" & GetFileSize(DirFiles.Size) & "<br>"
					Response.Write "�ļ����ͣ�" & DirFiles.Type & "<br>"
					Response.Write "�޸�ʱ�䣺" & DirFiles.DateLastModified
					FileCount = FileCount + 1
					Response.Write "</td>" & vbNewLine
					If (FileCount Mod 4) = 0 And FileCount < maxperpage And c < totalPut Then
						Response.Write "</tr>" & vbNewLine & "<tr>" & vbNewLine
					End If
				End If
				If FileCount >= maxperpage Then Exit For
			Next
			Response.Write "</tr>" & vbNewLine
			Response.Write "<tr><td colspan=4 class=tablerow1>" & vbNewLine
			Response.Write showpage(CurrentPage, totalPut, maxperpage, TotleSize)
			Response.Write "</td></tr>" & vbNewLine
			Response.Write "</table>"
		Else
			Response.Write "��Ŀ¼û���κ��ļ���"
		End If
		Set fsoFile = Nothing: Set fso = Nothing
	End Sub
	'=================================================
	'��������ShowChildFolder
	'��  �ã���ʾ��Ŀ¼�˵�
	'=================================================
	Private Sub ShowChildFolder()
		Dim fso, fsoFile, DirFolder
		Dim strFolderPath
		On Error Resume Next
		strFolderPath = ChannelDir & Request("UploadDir")
		strFolderPath = Server.MapPath(strFolderPath)
		Set fso = CreateObject(enchiasp.FSO_ScriptName)
		If fso.FolderExists(strFolderPath) Then
			Set fsoFile = fso.GetFolder(strFolderPath)
			For Each DirFolder In fsoFile.SubFolders
				Response.Write "<a href=""?ChannelID=" & ChannelID & "&UploadDir=" & Request("UploadDir") & "/" & DirFolder.Name& "&ThisDir=" & DirFolder.Name & """><img src=""images/pic/mediafolder.gif"" width=20 height=20 border=0 alt=""�޸�ʱ�䣺" & DirFolder.DateLastModified & """ align=absMiddle> "
				If Replace(ThisDir, "/", "") = DirFolder.Name Then
					Response.Write "<font color=red>" & DirFolder.Name & "</font>"
				Else
					Response.Write DirFolder.Name
				End If
				Response.Write "</a> &nbsp;&nbsp;" & vbNewLine
			Next
		Else
			Response.Write "û���ҵ��ļ��У�"
		End If
		Set fsoFile = Nothing: Set fso = Nothing
	End Sub

	'=================================================
	'��������showpage
	'��  �ã���ҳ
	'=================================================
	Private Function showpage(ByVal CurrentPage, ByVal TotalNumber, ByVal maxperpage, ByVal TotleSize)
		Dim n
		Dim strTemp
		
		If (TotalNumber Mod maxperpage) = 0 Then
			n = TotalNumber \ maxperpage
		Else
			n = TotalNumber \ maxperpage + 1
		End If
		strTemp = "<table align='center'><form method='Post' action='?ChannelID=" & ChannelID & "&UploadDir=" & Request("UploadDir") & "&ThisDir=" & Request("ThisDir") & "'><tr><td>" & vbNewLine
		strTemp = strTemp & "�� <b>" & TotalNumber & "</b> ���ļ���ռ�� <b>" & TotleSize & "</b>&nbsp;&nbsp;"
		'sfilename = JoinChar(sfilename)
		If CurrentPage < 2 Then
			strTemp = strTemp & "��ҳ ��һҳ&nbsp;"
		Else
			strTemp = strTemp & "<a href='?page=1&ChannelID=" & ChannelID & "&UploadDir=" & Request("UploadDir") & "&ThisDir=" & Request("ThisDir") & "'>��ҳ</a>&nbsp;"
			strTemp = strTemp & "<a href='?page=" & (CurrentPage - 1) & "&ChannelID=" & ChannelID & "&UploadDir=" & Request("UploadDir") & "&ThisDir=" & Request("ThisDir") & "'>��һҳ</a>&nbsp;"
		End If

		If n - CurrentPage < 1 Then
			strTemp = strTemp & "��һҳ βҳ"
		Else
			strTemp = strTemp & "<a href='?page=" & (CurrentPage + 1) & "&ChannelID=" & ChannelID & "&UploadDir=" & Request("UploadDir") & "&ThisDir=" & Request("ThisDir") & "'>��һҳ</a>&nbsp;"
			strTemp = strTemp & "<a href='?page=" & n & "&ChannelID=" & ChannelID & "&UploadDir=" & Request("UploadDir") & "&ThisDir=" & Request("ThisDir") & "'>βҳ</a>"
		End If
		strTemp = strTemp & "&nbsp;ҳ�Σ�<strong><font color=red>" & CurrentPage & "</font>/" & n & "</strong>ҳ "
		strTemp = strTemp & "&nbsp;ת����"
		strTemp = strTemp & "<input name=page size=3 value='" & CurrentPage & "'> <input type=submit name=Submit value='ת��' class=Button>"
		strTemp = strTemp & "</select>"
		strTemp = strTemp & "</td>"
		strTemp = strTemp & "<td>&nbsp;��<a href='#' onClick=""window.close();"">�رձ�����</a>��&nbsp;</td>"
		strTemp = strTemp & "</tr></form></table>"
		showpage = strTemp
	End Function
	'=================================================
	'��������GetFilePic
	'��  �ã���ȡ�ļ�ͼƬ
	'=================================================
	Private Function GetFilePic(sName)
		Dim FileName, Icon
		FileName = LCase(GetExtensionName(sName))
		Select Case FileName
			Case "gif", "jpg", "bmp", "png"
				Icon = sName
			Case "exe"
				Icon = "images/pic/file_exe.gif"
			Case "rar"
				Icon = "images/pic/file_rar.gif"
			Case "zip"
				Icon = "images/pic/file_zip.gif"
			Case "swf"
				Icon = "images/pic/file_flash.gif"
			Case "rm", "wma"
				Icon = "images/pic/file_rm.gif"
			Case "mid"
				Icon = "images/pic/file_media.gif"
			Case Else
				Icon = "images/pic/file_other.gif"
		End Select
		GetFilePic = Icon
	End Function
	'=================================================
	'��������GetExtensionName
	'��  �ã���ȡ�ļ���չ��
	'=================================================
	Private Function GetExtensionName(ByVal sName)
		Dim FileName
		FileName = Split(sName, ".")
		GetExtensionName = FileName(UBound(FileName))
	End Function
	'=================================================
	'��������GetFileSize
	'��  �ã���ʽ���ļ��Ĵ�С
	'=================================================
	Private Function GetFileSize(ByVal n)
		Dim FileSize
		FileSize = n / 1024
		FileSize = FormatNumber(FileSize, 2)
		If FileSize < 1024 And FileSize > 1 Then
			GetFileSize = "<font color=red>" & FileSize & "</font>&nbsp;KB"
		ElseIf FileSize > 1024 Then
			GetFileSize = "<font color=red>" & FormatNumber(FileSize / 1024, 2) & "</font>&nbsp;MB"
		Else
			GetFileSize = "<font color=red>" & n & "</font>&nbsp;Bytes"
		End If
	End Function
	'=================================================
	'��������DelFile
	'��  �ã�ɾ���ļ�
	'=================================================
	Private Sub DelFile()
		Dim fso, i
		Dim strFileName, strFilePath
		Dim strFolderName, strFolderPath
		'---- ɾ���ļ�
		If Trim(Request("FileName")) <> "" Then
			strFileName = Split(Request("FileName"), ",")
			If UBound(strFileName) <> -1 Then 'ɾ���ļ�
				Set fso = CreateObject(enchiasp.FSO_ScriptName)
				For i = 0 To UBound(strFileName)
					strFilePath = Server.MapPath(FilePath & Trim(strFileName(i)))
					If fso.FileExists(strFilePath) Then
						fso.DeleteFile strFilePath, True
					End If
				Next
				Set fso = Nothing
			End If
		End If
		'---- ɾ���ļ���
		If Trim(Request("FolderName")) <> "" Then
			strFolderName = Split(Request("FolderName"), ",")
			If UBound(strFolderName) <> -1 Then 'ɾ���ļ�
				Set fso = CreateObject(enchiasp.FSO_ScriptName)
				For i = 0 To UBound(strFolderName)
					strFolderPath = Server.MapPath(FilePath & Trim(strFolderName(i)))
					If fso.FolderExists(strFolderPath) Then
						fso.DeleteFolder strFolderPath, True
					End If
				Next
				Set fso = Nothing
			End If
		End If
		Response.Redirect (Request.ServerVariables("HTTP_REFERER"))
	End Sub
	'=================================================
	'��������DelAllDirFile
	'��  �ã�ɾ�������ļ����ļ���
	'=================================================
	Private Sub DelAllDirFile()
		Dim fso, oFolder
		Dim DirFile, DirFolder
		Dim tempPath
		
		Set fso = CreateObject(enchiasp.FSO_ScriptName)
		If fso.FolderExists(fullPath) Then
			Set oFolder = fso.GetFolder(fullPath)
			'---- ɾ�������ļ�
			For Each DirFile In oFolder.Files
				tempPath = fullPath & "\" & DirFile.Name
				fso.DeleteFile tempPath, True
			Next
			'---- ɾ��������Ŀ¼
			For Each DirFolder In oFolder.SubFolders
				tempPath = fullPath & "\" & DirFolder.Name
				fso.DeleteFolder tempPath, True
			Next
			Set oFolder = Nothing
		End If
		Set fso = Nothing
		Response.Redirect (Request.ServerVariables("HTTP_REFERER"))
	End Sub
	'=================================================
	'��������DelThisAllFile
	'��  �ã�ɾ����ǰĿ¼�����ļ�
	'=================================================
	Private Sub DelThisAllFile()
		Dim fso, oFolder
		Dim DirFiles
		Dim tempPath
		
		Set fso = CreateObject(enchiasp.FSO_ScriptName)
		If fso.FolderExists(fullPath) Then
			Set oFolder = fso.GetFolder(fullPath)
			'---- ɾ�������ļ�
			For Each DirFiles In oFolder.Files
				tempPath = fullPath & "\" & DirFiles.Name
				fso.DeleteFile tempPath, True
			Next
			Set oFolder = Nothing
		End If
		Set fso = Nothing
		Response.Redirect (Request.ServerVariables("HTTP_REFERER"))
	End Sub
	'=================================================
	'��������DelEmptyFolder
	'��  �ã�ɾ�����п��ļ���
	'=================================================
	Private Sub DelEmptyFolder()
		Dim fso, oFolder
		Dim DirFolder, tempPath
		
		Set fso = CreateObject(enchiasp.FSO_ScriptName)
		If fso.FolderExists(fullPath) Then
			Set oFolder = fso.GetFolder(fullPath)
			'---- ɾ�����п���Ŀ¼
			For Each DirFolder In oFolder.SubFolders
				If DirFolder.Size = 0 Then
					tempPath = fullPath & "\" & DirFolder.Name
					fso.DeleteFolder tempPath, True
				End If
			Next
			Set oFolder = Nothing
		End If
		Set fso = Nothing
		Response.Redirect (Request.ServerVariables("HTTP_REFERER"))
	End Sub
	'=================================================
	'��������ShowUploadMain
	'��  �ã���ʾ�ϴ��ļ���ҳ��
	'=================================================
	Private Sub ShowUploadMain()
		Dim maxperpage, CurrentPage, TotalNumber, Pcount
		Dim fso, FileCount, TotleSize, totalPut

		maxperpage = 20 '###ÿҳ��ʾ��
		
		If IsNumeric(Request("page")) And Trim(Request("page")) <> "" Then
			CurrentPage = CLng(Request("page"))
		Else
			CurrentPage = 1
		End If
		If CLng(CurrentPage) = 0 Then CurrentPage = 1
		On Error Resume Next
		If Not IsObjInstalled(enchiasp.FSO_ScriptName) Then
			Response.Write "<b><font color=red>��ķ�������֧�� fso(Scripting.FileSystemObject)! ����ʹ�ñ�����</font></b>"
		End If
		
		Response.Write "<table border=0 align=center cellpadding=3 cellspacing=1 class=tableborder>"
		Response.Write "<tr>"
		Response.Write "        <th colspan=""2"">�ļ�Ŀ¼</th>"
		Response.Write "</tr>"
		Response.Write "<tr>"
		Response.Write "        <td class=tablerow1 colspan=""2"">"
		Call ShowChildFolder
		Response.Write "</td>"
		Response.Write "</tr>"
		Response.Write "<tr>"
		Response.Write "        <td width=""50%"" class=tablerow2>��ǰĿ¼��" & FilePath & "</td>"
		Response.Write "        <td width=""50%"" align=center class=tablerow2>"
		Response.Write "<a href=""admin_UploadFile.asp?action=clear&ChannelID=" & ChannelID & "&UploadDir=" & Request("UploadDir") & """>���������ļ�</a> &nbsp;&nbsp;"
		If Trim(Request("ThisDir")) <> "" Then
			'Response.Write "<a href=""admin_UploadFile.asp?ChannelID=" & ChannelID & "&UploadDir=" & Request("UploadDir") & "&ThisDir=" & Left(Request("UploadDir"), InStrRev(Left(Request("ThisDir"), Len(Request("ThisDir")) - 1), "/")) & """>��������һ��Ŀ¼</a>"
			Response.Write "<a href=""admin_UploadFile.asp?ChannelID=" & ChannelID & "&UploadDir=" & Left(Request("UploadDir"),Len(Request("UploadDir"))-Len(Mid(Request("UploadDir"), InStrRev(Request("UploadDir"), "/")))) & "&ThisDir=" & Request("ThisDir") & """>��������һ��Ŀ¼</a>"
		End If
		Response.Write "</td>"
		Response.Write "</tr>"
		Response.Write "</table><br>" & vbNewLine

		Set fso = CreateObject(enchiasp.FSO_ScriptName)
		If fso.FolderExists(fullPath) Then
			Dim fsoFile, fsoFileSize
			Dim DirFiles, DirFolder
			Set fsoFile = fso.GetFolder(fullPath)
			'fsoFileSize = fsoFile.size '�ռ��Сͳ��
			Dim c
			FileCount = fsoFile.Files.Count
			TotleSize = GetFileSize(fsoFile.Size)
			totalPut = fsoFile.Files.Count
			If CurrentPage < 1 Then
				CurrentPage = 1
			End If
			If (CurrentPage - 1) * maxperpage > totalPut Then
				If (totalPut Mod maxperpage) = 0 Then
					CurrentPage = totalPut \ maxperpage
				Else
					CurrentPage = totalPut \ maxperpage + 1
				End If
			End If
			FileCount = 0
			c = 0
			Response.Write "<table border=0 align=center cellpadding=3 cellspacing=1 class=tableborder>" & vbNewLine
			Response.Write "<tr><td colspan=4 class=tablerow1>" & vbNewLine
			Response.Write showpage(CurrentPage, totalPut, maxperpage, TotleSize)
			Response.Write "</td></tr>" & vbNewLine
			Response.Write "<form name=""myform"" method=""post"" action='admin_uploadfile.asp'>" & vbCrLf
			Response.Write "<tr>" & vbNewLine
			Response.Write "<input type=hidden name=action value='del'>" & vbNewLine
			Response.Write "<input type=hidden name=ChannelID value='" & ChannelID & "'>" & vbNewLine
			Response.Write "<input type=hidden name=UploadDir value='" & Request("UploadDir") & "'>" & vbNewLine
			Response.Write "<input type=hidden name=ThisDir value='" & Request("ThisDir") & "'>" & vbNewLine
			For Each DirFiles In fsoFile.Files
				c = c + 1
				If c > maxperpage * (CurrentPage - 1) Then
					Response.Write "<td class=tablerow2>"
					Response.Write "<div align=center><a href='" & FilePath & DirFiles.Name & "'target=_blank><img src='" & GetFilePic(FilePath & DirFiles.Name) & "' width=140 height=100 border=0 alt='���ͼƬ�鿴ԭʼ�ļ���'></a></div>"
					Response.Write "�ļ�����<a href='" & FilePath & DirFiles.Name & "'target=_blank>" & DirFiles.Name & "</a><br>"
					Response.Write "�ļ���С��" & GetFileSize(DirFiles.Size) & "<br>"
					Response.Write "�ļ����ͣ�" & DirFiles.Type & "<br>"
					Response.Write "�޸�ʱ�䣺" & DirFiles.DateLastModified & "<br>"
					Response.Write "���������<input type=checkbox name=FileName value='" & DirFiles.Name & "' checked> ѡ��&nbsp;&nbsp;"
					Response.Write "<a href='?action=del&ChannelID=" & ChannelID & "&UploadDir=" & Request("UploadDir") & "&ThisDir=" & Request("ThisDir") & "&FileName=" & DirFiles.Name & "' onclick=""return confirm('��ȷ��Ҫɾ�����ļ���!');"">��ɾ��</a>"
					FileCount = FileCount + 1
					Response.Write "</td>" & vbNewLine
					If (FileCount Mod 4) = 0 And FileCount < maxperpage And c < totalPut Then
						Response.Write "</tr>" & vbNewLine & "<tr>" & vbNewLine
					End If
				End If
				If FileCount >= maxperpage Then Exit For
			Next
			Response.Write "</tr>" & vbNewLine
			Response.Write "<tr><td colspan=4 class=tablerow1>" & vbNewLine
			Response.Write "<input class=Button type=button name=chkall value='ȫѡ' onClick=""CheckAll(this.form)""><input class=Button type=button name=chksel value='��ѡ' onClick=""ContraSel(this.form)"">" & vbNewLine
			Response.Write "&nbsp;&nbsp;<input class=Button type=submit name=Submit2 value='ɾ��ѡ�е��ļ�' onClick=""return confirm('ȷ��Ҫɾ��ѡ�е��ļ���')"">" & vbNewLine
			Response.Write "&nbsp;&nbsp;<input class=Button type=submit name=Submit3 value='ɾ�������ļ�' onClick=""document.myform.action.value='DelThisAllFile';return confirm('ȷ��Ҫɾ����ǰĿ¼�����ļ���')"">" & vbNewLine
			Response.Write "&nbsp;&nbsp;<input class=Button type=submit name=Submit4 value='ɾ�������ļ����ļ���' onClick=""document.myform.action.value='DelAllDirFile';return confirm('ȷ��Ҫɾ����ǰĿ¼���ļ����ļ�����')"">" & vbNewLine
			Response.Write "&nbsp;&nbsp;<input class=Button type=submit name=Submit5 value='ɾ�����п��ļ���' onClick=""document.myform.action.value='DelEmptyFolder';return confirm('ȷ��Ҫɾ����ǰĿ¼���п��ļ�����')"">" & vbNewLine
			Response.Write "</tr></form>" & vbNewLine
			Response.Write "<tr><td colspan=4 class=tablerow1>" & vbNewLine
			Response.Write showpage(CurrentPage, totalPut, maxperpage, TotleSize)
			Response.Write "</td></tr>" & vbNewLine
			Response.Write "</table>"
		Else
			Response.Write "��Ŀ¼û���κ��ļ���"
		End If
		Set fsoFile = Nothing: Set fso = Nothing
	End Sub
	'=================================================
	'��������ClearUploadFile
	'��  �ã��������õ��ϴ��ļ�
	'=================================================
	Private Sub ClearUploadFile()
	
		Response.Write "<table border=0 align=center cellpadding=3 cellspacing=1 class=tableborder>" & vbNewLine
		Response.Write "<tr><th>" & vbNewLine
		If LCase(Request("UploadDir")) = "uploadfile" Then
			Response.Write "�������õ��ϴ��ļ�"
		Else
			Response.Write "�������õ��ϴ�ͼƬ"
		End If
		Response.Write "</th></tr>" & vbNewLine
		Response.Write "<form name=""myform"" method=""post"" action='admin_uploadfile.asp'>" & vbCrLf
		Response.Write "<input type=hidden name=action value='delete'>" & vbNewLine
		Response.Write "<input type=hidden name=ChannelID value='" & ChannelID & "'>" & vbNewLine
		Response.Write "<input type=hidden name=UploadDir value='" & Request("UploadDir") & "'>" & vbNewLine
		Response.Write "<tr><td class=tablerow1>" & vbNewLine
		Response.Write "<br>&nbsp;&nbsp;�١������վ��ʹ��һ��ʱ��󣬾ͻ�����������������ļ���������Ҫ����ʹ�ñ����ܽ�������<br>"
		Response.Write "<br>&nbsp;&nbsp;�ڡ���ȷ������ϴ�Ŀ¼��UploadPic��UploadFile����û��ʹ�õ��ļ����������ļ���<br>"
		Response.Write "<br>&nbsp;&nbsp;�ۡ�����ϴ��ļ��ܶ࣬�������ݿ����Ϣ���϶ִ࣬�б�������Ҫ�ķ��൱����ʱ�䣬���ڷ�������ʱִ�б�������<br>"
		Response.Write "<br></td></tr>" & vbNewLine
		Response.Write "<tr align=center><td class=tablerow2>��ѡ��Ҫ�����Ŀ¼��"
		Call ShowFolderPath
		Response.Write "<input class=Button type=submit name=Submit2 value=' ��ʼ���������ļ� ' onclick=""return confirm('��ȷ��Ҫ����������õ��ļ���');"">"
		Response.Write " ����<a href='?ChannelID=" & ChannelID & "&UploadDir=" & Request("UploadDir") & "'>�����ϴ�����</a>"
		Response.Write "</td></tr></form>" & vbNewLine
		Response.Write "</table>"
	End Sub
	'=================================================
	'��������ShowFolderPath
	'��  �ã���ʾ��Ŀ¼�˵�
	'=================================================
	Private Sub ShowFolderPath()
		Dim fso, fsoFile, DirFolder
		Dim strFolderPath
		On Error Resume Next
		strFolderPath = ChannelDir & Request("UploadDir")
		strFolderPath = Server.MapPath(strFolderPath)
		Set fso = CreateObject(enchiasp.FSO_ScriptName)
		If fso.FolderExists(strFolderPath) Then
			Set fsoFile = fso.GetFolder(strFolderPath)
			Response.Write "<select name=""path"">" & vbNewLine
			For Each DirFolder In fsoFile.SubFolders
				Response.Write "	<option value=""" & DirFolder.Name & """>" & DirFolder.Name & "</option>" & vbNewLine
			Next
			Response.Write "	<option value="""">�ϴ���Ŀ¼</option>" & vbNewLine
			Response.Write "</select>" & vbNewLine
			Set fsoFile = Nothing
		Else
			'Response.Write "û���ҵ��ļ��У�"
		End If
		Set fso = Nothing
	End Sub
	'=================================================
	'��������DelUselessFile
	'��  �ã�ɾ���������õ��ϴ��ļ�
	'=================================================
	Private Sub DelUselessFile()
		Dim SQL,Rs,i
		Dim fso, fsoFile, DirFiles
		Dim strFileName,strFolderPath
		Dim strFilePath,strDirName
		Server.ScriptTimeout = 9999999
		On Error Resume Next
		
		If Len(Request("path")) > 0 Then
			strDirName = Request("path") & "/"
		Else
			strDirName = vbNullString
		End If
		strFolderPath = ChannelDir & UploadDir & strDirName
		strFolderPath = Server.MapPath(strFolderPath)
		Set fso = CreateObject(enchiasp.FSO_ScriptName)
		i = 0
		If fso.FolderExists(strFolderPath) Then
			Set fsoFile = fso.GetFolder(strFolderPath)
			For Each DirFiles In fsoFile.Files
			
				strFileName = strDirName & DirFiles.Name
				strFilePath = strFolderPath & "\" & DirFiles.Name
				Select Case CLng(modules)
				case -1
				Case 1
					SQL = "SELECT TOP 1 ArticleID FROM [ECCMS_Article] WHERE ChannelID=" & ChannelID & " And UploadImage like '%" & strFileName & "%'"
				Case 2
					If LCase(Request("UploadDir")) = "uploadfile" Then
						'SQL = "SELECT TOP 1 softid FROM [ECCMS_SoftList] WHERE ChannelID=" & ChannelID & " And DownAddress like '%" & strFileName & "%'"
						SQL = "SELECT TOP 1 id FROM [ECCMS_DownAddress] WHERE ChannelID=" & ChannelID & " And DownFileName like '%" & strFileName & "%'"
					Else
						SQL = "SELECT TOP 1 softid FROM [ECCMS_SoftList] WHERE ChannelID=" & ChannelID & " And SoftImage like '%" & strFileName & "%'"
					End If
				Case 5
					If LCase(Request("UploadDir")) = "uploadfile" Then
						SQL = "SELECT TOP 1 flashid FROM [ECCMS_FlashList] WHERE ChannelID=" & ChannelID & " And showurl like '%" & strFileName & "%'"
					Else
						SQL = "SELECT TOP 1 flashid FROM [ECCMS_FlashList] WHERE ChannelID=" & ChannelID & " And miniature like '%" & strFileName & "%'"
					End If
				case 6
					SQL = "SELECT TOP 1 ArticleID FROM [ECCMS_Article] WHERE ChannelID=" & ChannelID & " And UploadImage like '%" & strFileName & "%'"

				Case Else
					SQL = "SELECT TOP 1 id FROM [ECCMS_Adlist] WHERE Picurl like '%" & strFileName & "%'"
				End Select
				Set Rs = enchiasp.Execute(SQL)
				If Rs.EOF Then
					i = i + 1
					fso.DeleteFile(strFilePath)
				End If
			Next
			Set fsoFile = Nothing
		End If
		Set fso = Nothing
		Succeed ("<li>�ļ�������ɣ�</li><li>һ��������<font color=red><b>" & i & "</b></font>�������ļ�")
	End Sub
	
End Class
%>
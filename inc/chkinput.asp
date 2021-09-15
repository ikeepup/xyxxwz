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
'判断EMAIL是否正确
Function IsValidEmail(email)
	Dim names, Name, i, c
	IsValidEmail = true
	names = Split(email, "@")
	If UBound(names) <> 1 Then
		IsValidEmail = false
		Exit Function
	End If
	For Each Name in names
		If Len(Name) <= 0 Then
			IsValidEmail = false
			Exit Function
		End If
		For i = 1 To Len(Name)
			c = LCase(Mid(Name, i, 1))
			If InStr("abcdefghijklmnopqrstuvwxyz_-.", c) <= 0 And Not IsNumeric(c) Then
				IsValidEmail = false
				Exit Function
			End If
		Next
		If Left(Name, 1) = "." Or Right(Name, 1) = "." Then
			IsValidEmail = false
			Exit Function
		End If
	Next
	If InStr(names(1), ".") <= 0 Then
		IsValidEmail = false
		Exit Function
	End If
	i = Len(names(1)) - InStrRev(names(1), ".")
	If i <> 2 And i <> 3 Then
		IsValidEmail = false
		Exit Function
	End If
	If InStr(email, "..") > 0 Then
		IsValidEmail = false
	End If

End Function
' 判断电话号码是否正确
Function IsValidTel(para)
	Dim Str
	Dim l, i
	If IsNull(para) Then
		IsValidTel = false
		Exit Function
	End If
	Str = CStr(para)
	If Len(Trim(Str))<7 Then
		IsValidTel = false
		Exit Function
	End If
	l = Len(Str)
	For i = 1 To l
		If Not (Mid(Str, i, 1)>= "0" And Mid(Str, i, 1)<= "9" Or Mid(Str, i, 1) = "-") Then
			IsValidTel = false
			Exit Function
		End If
	Next
	IsValidTel = true
End Function

public function wordhelp()
	response.write "<html>"
	response.write "<head>"
	response.write "<tittle>"
	response.write "</tittle>"
	response.write "<body>"
	response.write "</body>"
	end function

%>

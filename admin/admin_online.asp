<!--#include file="setup.asp"-->
<!--#include file="check.asp"-->
<%
Admin_header
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
%>
<table border=0 align=center cellpadding=3 cellspacing=1 class=tableborder>
<tr>
	<th>����������Ϣ����</th>
</tr>
<tr>
	<td class=tablerow1>�˵�������<a href='admin_online.asp'>������ҳ</a> | 
	<a href='admin_online.asp?action=zone'>��ϸ��ַ</a> | 
	<a href='admin_online.asp?action=refer'>������Դ</a> |
	<a href='admin_online.asp?action=delall' onclick="{if(confirm('��ȷ��Ҫɾ����������������?')){return true;}return false;}"><font color=blue>ɾ��������������</font></a></td>
</tr>
<tr>
	<td class=tablerow2>��ǰλ�ã���������ͳ����Ϣ</td>
</tr>
</table>
<br>
<%
Dim Action,i,strClass,sFileName
Dim maxperpage,CurrentPage,totalnumber,TotalPageNum

maxperpage = 30 '###ÿҳ��ʾ��
If Not IsNumeric(Request("page")) And Len(Request("page")) <> 0 Then
	Response.Write ("�����ϵͳ����!����������")
	Response.End
End If
If Not IsEmpty(Request("page")) And Len(Request("page")) <> 0 Then
	CurrentPage = CInt(Request("page"))
Else
	CurrentPage = 1
End If
If CInt(CurrentPage) = 0 Then CurrentPage = 1
TotalNumber = enchiasp.Execute("Select Count(ID) from ECCMS_Online")(0)
TotalPageNum = CInt(TotalNumber / maxperpage)  '�õ���ҳ��
If TotalPageNum < TotalNumber / maxperpage Then TotalPageNum = TotalPageNum + 1
If CurrentPage < 1 Then CurrentPage = 1
If CurrentPage > TotalPageNum Then CurrentPage = TotalPageNum

Action = LCase(Request("action"))
If Not ChkAdmin("Online") Then
	Server.Transfer("showerr.asp")
	Response.End
End If
Select Case Trim(Action)
	Case "refer"
		Call OnlineReferer
	Case "zone"
		Call OnlineZone
	Case "del"
		Call DelOnline
	Case "delall"
		Call DelAllOnline
	Case Else
		Call MainPage
End Select
If FoundErr = True Then
	ReturnError(ErrMsg)
End If
Admin_footer
SaveLogInfo(AdminName)
CloseConn
Private Sub MainPage()
%>
<table border=0 align=center cellpadding=3 cellspacing=1 class=tableborder>
<tr>
	<th width='5%' nowrap>ѡ��</th>
	<th nowrap>�� �� ��</th>
	<th nowrap>�� �� ʱ ��</th>
	<th nowrap>�� �� ʱ ��</th>
	<th nowrap>�� �� IP �� ַ</th>
	<th nowrap>�� �� ϵ ͳ</th>
	<th nowrap>� �� ��</th>
</tr>
<%
	sFileName = "admin_online.asp?"
	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL = "select * from [ECCMS_Online] order by startTime desc"
	If IsSqlDataBase = 1 Then
		If CurrentPage > 100 Then
			Rs.Open SQL, Conn, 1, 1
		Else
			Set Rs = Conn.Execute(SQL)
		End If
	Else
		Rs.Open SQL, Conn, 1, 1
	End If
	enchiasp.SqlQueryNum = enchiasp.SqlQueryNum + 1
	If Rs.bof And Rs.EOF Then
		Response.Write "<tr><td align=center colspan=7 class=TableRow2>��ǰ�������ߣ�</td></tr>"
	Else
		If TotalPageNum > 1 then Rs.Move (CurrentPage - 1) * maxperpage
		i = 0
%>
<tr>
	<td class=tablerow2 colspan=7><%Call showpage()%></td>
</tr>
<form name=selform method=post action=?action=del>
<%
	Do While Not Rs.EOF And i < CInt(maxperpage)
		If Not Response.IsClientConnected Then Response.End
		If (i mod 2) = 0 Then
			strClass = "class=TableRow1"
		Else
			strClass = "class=TableRow2"
		End If
%>
<tr align=center>
	<td <%=strClass%>><input type=checkbox name=OnlineID value='<%=Rs("id")%>'></td>
	<td <%=strClass%>><%=Rs("username")%></td>
	<td <%=strClass%>><%=Rs("startTime")%></td>
	<td <%=strClass%>><%=Rs("lastTime")%></td>
	<td <%=strClass%>><%=Rs("ip")%></td>
	<td <%=strClass%>><%=usersysinfo(Rs("browser"), 0)%></td>
	<td <%=strClass%>><%=usersysinfo(Rs("browser"), 1)%></td>
</tr>
<%
		Rs.movenext
			i = i + 1
			If i >= maxperpage Then Exit Do
		Loop
	End If
	Rs.Close:Set Rs = Nothing
%>
<tr>
	<td class=tablerow1 colspan=7>
	<input class=Button type="button" name="chkall" value="ȫѡ" onClick="CheckAll(this.form)"><input class=Button type="button" name="chksel" value="��ѡ" onClick="ContraSel(this.form)">
	<input class=Button type="submit" name="Submit2" value="ɾ��" onclick="{if(confirm('��ȷ��Ҫɾ����������Ա��?')){this.document.selform.submit();return true;}return false;}">
	<input class=Button type="button" name="Submit3" value="���������������" onclick="{if(confirm('��ȷ��Ҫ�����������������?')){location.href='admin_online.asp?action=delall';return true;}return false;}"></td>
</tr>
</form>
<tr>
	<td class=tablerow2 colspan=7><%Call showpage()%></td>
</tr>
</table>

<%
End Sub

Private Sub OnlineReferer()
%>
<table border=0 align=center cellpadding=3 cellspacing=1 class=tableborder>
<tr>
	<th width='5%' nowrap>ѡ��</th>
	<th width='15%' nowrap>����ʱ��/IP</th>
	<th>�� �� �� Դ</th>
	<th>�� ǰ λ ��</th>
</tr>
<%
	sFileName = "admin_online.asp?action=refer&"
	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL = "select * from [ECCMS_Online] order by startTime desc"
	If IsSqlDataBase = 1 Then
		If CurrentPage > 100 Then
			Rs.Open SQL, Conn, 1
		Else
			Set Rs = Conn.Execute(SQL)
		End If
	Else
		Rs.Open SQL, Conn, 1
	End If
	enchiasp.SqlQueryNum = enchiasp.SqlQueryNum + 1
	If Rs.bof And Rs.EOF Then
		Response.Write "<tr><td align=center colspan=4 class=TableRow2>��ǰ�������ߣ�</td></tr>"
	Else
		If TotalPageNum > 1 then Rs.Move (CurrentPage - 1) * maxperpage
		i = 0
%>
<tr>
	<td class=tablerow2 colspan=4><%Call showpage()%></td>
</tr>
<form name=selform method=post action=?action=del>
<%
	Do While Not Rs.EOF And i < CInt(maxperpage)
		If Not Response.IsClientConnected Then Response.End
		If (i mod 2) = 0 Then
			strClass = "class=TableRow1"
		Else
			strClass = "class=TableRow2"
		End If
%>
<tr>
	<td align=center <%=strClass%>><input type=checkbox name=OnlineID value='<%=Rs("id")%>'></td>
	<td align=center <%=strClass%> nowrap><%=Rs("startTime")%><br><%=Rs("ip")%></td>
	<td <%=strClass%>><a href='<%=Rs("strReferer")%>' target=_blank><%=Rs("strReferer")%></a></td>
	<td <%=strClass%>><a href='<%=Rs("station")%>' target=_blank><%=Rs("station")%></a></td>
</tr>
<%
		Rs.movenext
			i = i + 1
			If i >= maxperpage Then Exit Do
		Loop
	End If
	Rs.Close:Set Rs = Nothing
%>
<tr>
	<td class=tablerow1 colspan=4>
	<input class=Button type="button" name="chkall" value="ȫѡ" onClick="CheckAll(this.form)"><input class=Button type="button" name="chksel" value="��ѡ" onClick="ContraSel(this.form)">
	<input class=Button type="submit" name="Submit2" value="ɾ��" onclick="{if(confirm('��ȷ��Ҫɾ����������Ա��?')){this.document.selform.submit();return true;}return false;}">
	<input class=Button type="button" name="Submit3" value="���������������" onclick="{if(confirm('��ȷ��Ҫ�����������������?')){location.href='admin_online.asp?action=delall';return true;}return false;}"></td>
</tr>
</form>
<tr>
	<td class=tablerow2 colspan=7><%Call showpage()%></td>
</tr>
</table>

<%
End Sub

Private Sub OnlineZone()
%>
<table border=0 align=center cellpadding=3 cellspacing=1 class=tableborder>
<tr>
	<th width='5%' nowrap>ѡ��</th>
	<th nowrap>�� �� ��</th>
	<th nowrap>�� �� ��</th>
	<th nowrap>IP �� ַ</th>
	<th nowrap>�� ϸ �� ַ</th>
	<th nowrap>�� �� ϵ ͳ</th>
	<th nowrap>� �� ��</th>
</tr>
<%
	sFileName = "admin_online.asp?action=zone&"
	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL = "select * from [ECCMS_Online] order by startTime desc"
	If IsSqlDataBase = 1 Then
		If CurrentPage > 100 Then
			Rs.Open SQL, Conn, 1
		Else
			Set Rs = Conn.Execute(SQL)
		End If
	Else
		Rs.Open SQL, Conn, 1
	End If
	enchiasp.SqlQueryNum = enchiasp.SqlQueryNum + 1
	If Rs.bof And Rs.EOF Then
		Response.Write "<tr><td align=center colspan=7 class=TableRow2>��ǰ�������ߣ�</td></tr>"
	Else
		If TotalPageNum > 1 then Rs.Move (CurrentPage - 1) * maxperpage
		i = 0
%>
<tr>
	<td class=tablerow2 colspan=7><%Call showpage()%></td>
</tr>
<form name=selform method=post action=?action=del>
<%
	Do While Not Rs.EOF And i < CInt(maxperpage)
		If Not Response.IsClientConnected Then Response.End
		If (i mod 2) = 0 Then
			strClass = "class=TableRow1"
		Else
			strClass = "class=TableRow2"
		End If
%>
<tr align=center>
	<td <%=strClass%>><input type=checkbox name=OnlineID value='<%=Rs("id")%>'></td>
	<td <%=strClass%>><%=Rs("username")%></td>
	<td <%=strClass%>><%=Rs("identitys")%></td>
	<td <%=strClass%>><%=Rs("ip")%></td>
	<td <%=strClass%>><%=GetAddress(Rs("ip"))%></td>
	<td <%=strClass%>><%=usersysinfo(Rs("browser"), 0)%></td>
	<td <%=strClass%>><%=usersysinfo(Rs("browser"), 1)%></td>
</tr>
<%
		Rs.movenext
			i = i + 1
			If i >= maxperpage Then Exit Do
		Loop
	End If
	Rs.Close:Set Rs = Nothing
%>
<tr>
	<td class=tablerow1 colspan=7>
	<input class=Button type="button" name="chkall" value="ȫѡ" onClick="CheckAll(this.form)"><input class=Button type="button" name="chksel" value="��ѡ" onClick="ContraSel(this.form)">
	<input class=Button type="submit" name="Submit2" value="ɾ��" onclick="{if(confirm('��ȷ��Ҫɾ����������Ա��?')){this.document.selform.submit();return true;}return false;}">
	<input class=Button type="button" name="Submit3" value="���������������" onclick="{if(confirm('��ȷ��Ҫ�����������������?')){location.href='admin_online.asp?action=delall';return true;}return false;}"></td>
</tr>
</form>
<tr>
	<td class=tablerow2 colspan=7><%Call showpage()%></td>
</tr>
</table>

<%
End Sub

Private Sub DelAllOnline()
	enchiasp.Execute("Delete From ECCMS_Online")
	Call OutputScript ("��������ȫ�������ɣ�","admin_online.asp")
End Sub

Private Sub DelOnline()
	Dim OnlineID
	If Request("OnlineID") <> "" Then
		OnlineID = Request("OnlineID")
		enchiasp.Execute("Delete From ECCMS_Online where ID in (" & OnlineID & ")")
		OutHintScript ("��������ɾ���ɹ���")
	Else
		OutAlertScript("��ѡ����ȷ��ϵͳ������")
	End If
End Sub
Private Sub showpage()
	Dim n
	If totalnumber Mod maxperpage = 0 Then
		n = totalnumber \ maxperpage
	Else
		n = totalnumber \ maxperpage + 1
	End If
	Response.Write "<table cellspacing=1 width='100%' border=0><form method=Post action='" & sFileName & "'><tr><td align=center> " & vbCrLf
	If CurrentPage < 2 Then
		Response.Write "��ǰ�������� <font COLOR=#FF0000><strong>" & totalnumber & "</strong></font> ��&nbsp;�� ҳ&nbsp;��һҳ&nbsp;|&nbsp;"
	Else
		Response.Write "��ǰ�������� <font COLOR=#FF0000><strong>" & totalnumber & "</strong></font> ��&nbsp;<a href=" & sFileName & "page=1>�� ҳ</a>&nbsp;"
		Response.Write "<a href=" & sFileName & "page=" & CurrentPage - 1 & ">��һҳ</a>&nbsp;|&nbsp;"
	End If
	If n - CurrentPage < 1 Then
		Response.Write "��һҳ&nbsp;β ҳ" & vbCrLf
	Else
		Response.Write "<a href=" & sFileName & "page=" & (CurrentPage + 1) & ">��һҳ</a>"
		Response.Write "&nbsp;<a href=" & sFileName & "page=" & n & ">β ҳ</a>" & vbCrLf
	End If
	Response.Write "&nbsp;ҳ�Σ�<strong><font color=red>" & CurrentPage & "</font>/" & n & "</strong>ҳ "
	Response.Write "&nbsp;ת����"
	Response.Write "<input name=page size=3 value='" & CurrentPage & "'> <input class=Button type=submit name=Submit value='ת��'>"
	Response.Write "</td></tr></FORM></table>" & vbCrLf
End Sub

Private Function usersysinfo(info, getinfo)
	Dim usersys
	usersys = Split(info, "|")
	usersysinfo = usersys(getinfo)
End Function

Public Function GetAddress(sip)
	If Len(sip) < 5 Then
		GetAddress = "δ֪"
		Exit Function
	End If
	On Error Resume Next
	Dim Wry,IPType
	Set Wry = New TQQWry
	If Not Wry.IsIp(sip) Then
		GetAddress = " δ֪"
		Exit Function
	End If
	IPType = Wry.QQWry(sip)
	GetAddress = Wry.Country & " " & Wry.LocalStr
End Function

Class TQQWry
	' ============================================
	' ��������
	' ============================================
	Dim Country, LocalStr, Buf, OffSet
	Private StartIP, EndIP, CountryFlag
	Public QQWryFile
	Public FirstStartIP, LastStartIP, RecordCount
	Private Stream, EndIPOff
	' ============================================
	' ��ģ���ʼ��
	' ============================================
	Private Sub Class_Initialize
		On Error Resume Next
		Country 		= ""
		LocalStr 		= ""
		StartIP 		= 0
		EndIP 			= 0
		CountryFlag 	= 0 
		FirstStartIP 	= 0 
		LastStartIP 	= 0 
		EndIPOff 		= 0 
		QQWryFile = Server.MapPath("../DataBase/IPAddress.dat") 'QQ IP��·����Ҫת��������·��
	End Sub
	' ============================================
	' IP��ַת��������
	' ============================================
	Function IPToInt(IP)
		Dim IPArray, i
		IPArray = Split(IP, ".", -1)
		FOr i = 0 to 3
			If Not IsNumeric(IPArray(i)) Then IPArray(i) = 0
			If CInt(IPArray(i)) < 0 Then IPArray(i) = Abs(CInt(IPArray(i)))
			If CInt(IPArray(i)) > 255 Then IPArray(i) = 255
		Next
		IPToInt = (CInt(IPArray(0))*256*256*256) + (CInt(IPArray(1))*256*256) + (CInt(IPArray(2))*256) + CInt(IPArray(3))
	End Function
	' ============================================
	' ������תIP��ַ
	' ============================================
	Function IntToIP(IntValue)
		p4 = IntValue - Fix(IntValue/256)*256
		IntValue = (IntValue-p4)/256
		p3 = IntValue - Fix(IntValue/256)*256
		IntValue = (IntValue-p3)/256
		p2 = IntValue - Fix(IntValue/256)*256
		IntValue = (IntValue - p2)/256
		p1 = IntValue
		IntToIP = Cstr(p1) & "." & Cstr(p2) & "." & Cstr(p3) & "." & Cstr(p4)
	End Function
	' ============================================
	' ��ȡ��ʼIPλ��
	' ============================================
	Private Function GetStartIP(RecNo)
		OffSet = FirstStartIP + RecNo * 7
		Stream.Position = OffSet
		Buf = Stream.Read(7)
		
		EndIPOff = AscB(MidB(Buf, 5, 1)) + (AscB(MidB(Buf, 6, 1))*256) + (AscB(MidB(Buf, 7, 1))*256*256) 
		StartIP  = AscB(MidB(Buf, 1, 1)) + (AscB(MidB(Buf, 2, 1))*256) + (AscB(MidB(Buf, 3, 1))*256*256) + (AscB(MidB(Buf, 4, 1))*256*256*256)
		GetStartIP = StartIP
	End Function
	' ============================================
	' ��ȡ����IPλ��
	' ============================================
	Private Function GetEndIP()
		Stream.Position = EndIPOff
		Buf = Stream.Read(5)
		EndIP = AscB(MidB(Buf, 1, 1)) + (AscB(MidB(Buf, 2, 1))*256) + (AscB(MidB(Buf, 3, 1))*256*256) + (AscB(MidB(Buf, 4, 1))*256*256*256) 
		CountryFlag = AscB(MidB(Buf, 5, 1))
		GetEndIP = EndIP
	End Function
	' ============================================
	' ��ȡ������Ϣ���������Һͺ�ʡ��
	' ============================================
	Private Sub GetCountry(IP)
		If (CountryFlag = 1 Or CountryFlag = 2) Then
			Country = GetFlagStr(EndIPOff + 4)
			If CountryFlag = 1 Then
				LocalStr = GetFlagStr(Stream.Position)
				' ����������ȡ���ݿ�汾��Ϣ
				If IP >= IPToInt("255.255.255.0") And IP <= IPToInt("255.255.255.255") Then
					LocalStr = GetFlagStr(EndIPOff + 21)
					Country = GetFlagStr(EndIPOff + 12)
				End If
			Else
				LocalStr = GetFlagStr(EndIPOff + 8)
			End If
		Else
			Country = GetFlagStr(EndIPOff + 4)
			LocalStr = GetFlagStr(Stream.Position)
		End If
		' �������ݿ��е�������Ϣ
		Country = Trim(Country)
		LocalStr = Trim(LocalStr)
		If InStr(Country, "CZ88.NET") Then Country = "GZ110.CN"
		If InStr(LocalStr, "CZ88.NET") Then LocalStr = "GZ110.CN"
	End Sub
	' ============================================
	' ��ȡIP��ַ��ʶ��
	' ============================================
	Private Function GetFlagStr(OffSet)
		Dim Flag
		Flag = 0
		Do While (True)
			Stream.Position = OffSet
			Flag = AscB(Stream.Read(1))
			If(Flag = 1 Or Flag = 2 ) Then
				Buf = Stream.Read(3) 
				If (Flag = 2 ) Then
					CountryFlag = 2
					EndIPOff = OffSet - 4
				End If
				OffSet = AscB(MidB(Buf, 1, 1)) + (AscB(MidB(Buf, 2, 1))*256) + (AscB(MidB(Buf, 3, 1))*256*256)
			Else
				Exit Do
			End If
		Loop
		
		If (OffSet < 12 ) Then
			GetFlagStr = ""
		Else
			Stream.Position = OffSet
			GetFlagStr = GetStr() 
		End If
	End Function
	' ============================================
	' ��ȡ�ִ���Ϣ
	' ============================================
	Private Function GetStr() 
		Dim c
		GetStr = ""
		Do While (True)
			c = AscB(Stream.Read(1))
			If (c = 0) Then Exit Do 
			
			'�����˫�ֽڣ��ͽ��и��ֽ��ڽ�ϵ��ֽںϳ�һ���ַ�
			If c > 127 Then
				If Stream.EOS Then Exit Do
				GetStr = GetStr & Chr(AscW(ChrB(AscB(Stream.Read(1))) & ChrB(C)))
			Else
				GetStr = GetStr & Chr(c)
			End If
		Loop 
	End Function
	' ============================================
	' ���ĺ�����ִ��IP����
	' ============================================
	Public Function QQWry(DotIP)
		Dim IP, nRet
		Dim RangB, RangE, RecNo
		
		IP = IPToInt (DotIP)
		
		Set Stream = CreateObject("ADodb.Stream")
		Stream.Mode = 3
		Stream.Type = 1
		Stream.Open
		Stream.LoadFromFile QQWryFile
		Stream.Position = 0
		Buf = Stream.Read(8)
		
		FirstStartIP = AscB(MidB(Buf, 1, 1)) + (AscB(MidB(Buf, 2, 1))*256) + (AscB(MidB(Buf, 3, 1))*256*256) + (AscB(MidB(Buf, 4, 1))*256*256*256)
		LastStartIP  = AscB(MidB(Buf, 5, 1)) + (AscB(MidB(Buf, 6, 1))*256) + (AscB(MidB(Buf, 7, 1))*256*256) + (AscB(MidB(Buf, 8, 1))*256*256*256)
		RecordCount = Int((LastStartIP - FirstStartIP)/7)
		' �����ݿ����Ҳ����κ�IP��ַ
		If (RecordCount <= 1) Then
			Country = "δ֪"
			QQWry = 2
			Exit Function
		End If
		
		RangB = 0
		RangE = RecordCount
		
		Do While (RangB < (RangE - 1)) 
			RecNo = Int((RangB + RangE)/2) 
			Call GetStartIP (RecNo)
			If (IP = StartIP) Then
				RangB = RecNo
				Exit Do
			End If
			If (IP > StartIP) Then
				RangB = RecNo
			Else 
				RangE = RecNo
			End If
		Loop
		
		Call GetStartIP(RangB)
		Call GetEndIP()

		If (StartIP <= IP) And ( EndIP >= IP) Then
			' û���ҵ�
			nRet = 0
		Else
			' ����
			nRet = 3
		End If
		Call GetCountry(IP)

		QQWry = nRet
	End Function
	' ============================================
	' ���IP��ַ�Ϸ���
	' ============================================
	Public Function IsIp(IP)
		IsIp = True
		If IP = "" Then IsIp = False : Exit Function
		Dim Re
		Set Re = New RegExp
		Re.Pattern = "^(0|[1-9]\d?|[0-1]\d{2}|2[0-4]\d|25[0-5])\.(0|[1-9]\d?|[0-1]\d{2}|2[0-4]\d|25[0-5])\.(0|[1-9]\d?|[0-1]\d{2}|2[0-4]\d|25[0-5])\.(0|[1-9]\d?|[0-1]\d{2}|2[0-4]\d|25[0-5])$"
		Re.IgnoreCase = True
		Re.Global = True
		IsIp = Re.Test(IP)
		Set Re = Nothing
	End Function
	' ============================================
	' ���ս�

	' ============================================
	Private Sub Class_Terminate
		On ErrOr Resume Next
		Stream.Close
		If Err Then Err.Clear
		Set Stream = Nothing
	End Sub
End Class 

%>
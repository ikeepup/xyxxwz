<!--#include file="config.asp"-->
<!--#include file="check.asp"-->
<!--#include file="head.inc"-->
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
Dim Rs,SQL,i,Action
Dim maxperpage,CurrentPage,totalnumber,TotalPageNum

Call InnerLocation("�ҵ��ղؼ�")

If CInt(GroupSetting(3)) = 0 Then
	ErrMsg = ErrMsg + "<li>�Բ�����û��ʹ���ղؼе�Ȩ�ޣ�����ʲô��������ϵ����Ա��</li>"
	Founderr = True
End If
Action = enchiasp.CheckStr(LCase(Trim(Request("action"))))
Select Case Trim(Action)
	Case "save","���"
		Call SaveFavorite
	Case "add"
		Call AddFavorite
	Case "del"
		Call DelFavorite
	Case "����ղؼ�"
		Call DelAllFavorite
	Case Else
		Call showmain
End Select

If Founderr = True Then
	Call Returnerr(ErrMsg)
End If
Sub showmain()
	If Founderr = True Then Exit Sub
	maxperpage = 20 '###ÿҳ��ʾ��
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
%>
<table cellspacing=1 align=center cellpadding=3 border=0 class=Usertableborder>
	<tr>
		<th colspan=3>>> �ҵ��ղؼ� <<</th>
	</tr>
	<tr>
		<td width="65%" align=center class=Usertablerow2><b class=userfont2>����</b></td>
		<td width="23%" align=center class=Usertablerow2><b class=userfont2>�ղ�ʱ��</b></td>
		<td width="12%" align=center class=Usertablerow2><b class=userfont2>����</b></td>
	</tr>
<%
	TotalNumber = enchiasp.Execute("Select Count(FavoriteID) from ECCMS_Favorite where username='"& enchiasp.CheckStr(enchiasp.membername) &"'")(0)
	TotalPageNum = CInt(TotalNumber / maxperpage)  '�õ���ҳ��
	If TotalPageNum < TotalNumber / maxperpage Then TotalPageNum = TotalPageNum + 1
	If CurrentPage < 1 Then CurrentPage = 1
	If CurrentPage > TotalPageNum Then CurrentPage = TotalPageNum
	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL = "select * from [ECCMS_Favorite] where username='"& enchiasp.CheckStr(enchiasp.membername) &"' order by FavoriteID desc"
	Rs.Open SQL, Conn, 1, 1
	enchiasp.SqlQueryNum = enchiasp.SqlQueryNum + 1
	If Not (Rs.bof And Rs.EOF) Then
		If TotalPageNum > 1 then Rs.Move (CurrentPage - 1) * maxperpage
		i = 0
		Do While Not Rs.EOF And i < CInt(maxperpage)
%>
	<tr>
		<td class=Usertablerow1><a href="<%=Rs("fondurl")%>" target=_blank><%=Server.HTMLEncode(Rs("fondtopic"))%></a></td>
		<td align=center class=Usertablerow1><%=Rs("addTime")%></td>
		<td align=center class=Usertablerow1><a href="?action=del&favid=<%=Rs("FavoriteID")%>" onclick="showClick('ɾ���󽫲��ָܻ�����ȷ��Ҫɾ����?')"><img src="images/delete.gif" width="52" height="16" border=0 alt="ɾ��"></a></td>
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
		<td colspan=3 align=center class=Usertablerow1><%Response.Write ShowPages (CurrentPage,TotalPageNum,TotalNumber,maxperpage,"")%></td>
	</tr>
	<tr>
		<th colspan=3>>> ����ղ� <<</th>
	</tr>
	<form name=myform method=post action="">
	<tr>
		<td colspan=3 align=center class=Usertablerow1><b class=userfont2>���⣺</b><input type="text" name="fondtopic" size=20>
		<b class=userfont2>URL��</b><input type="text" name="fondurl" size=30 value="http://">
		<input type=submit name="action" value="���" class=button> <input type=submit name="action" value="����ղؼ�" onclick="{if(confirm('��պ󽫲��ָܻ���ȷ��������еļ�¼��?')){this.document.myform.submit();return true;}return false;}" class=button><br>
		<div><b>ע�⣺</b><%If CLng(GroupSetting(5)) <> 0 Then%>�����ֻ���ղ� <b class=userfont1><%=GroupSetting(5)%></b> ����Ϣ��<%End If%>�붨ʱɾ�����õ���Ϣ��</div></td>
	</tr>
	</form>
</table>
<%
End Sub
'================================================
' ��������DelFavorite
' ��  �ã�ɾ���ղ���Ϣ
'================================================
Sub DelFavorite()
	If enchiasp.CheckPost=False Then
		ErrMsg = Postmsg
		Founderr = True
		Exit Sub
	End If
	If Not IsNumeric(Request("favid")) Then
		ErrMsg = ErrMsg + "<li>�Բ�����û��ʹ���ղؼе�Ȩ�ޣ�����ʲô��������ϵ����Ա��</li>"
		Founderr = True
		Exit Sub
	End If
	enchiasp.Execute("Delete From ECCMS_Favorite where username='"& enchiasp.membername &"' And FavoriteID="& CLng(Request("favid")))
	Call Returnsuc("<li>��¼ɾ���ɹ���</li>")
End Sub
'================================================
' ��������DelAllFavorite
' ��  �ã�����û��ղؼ�
'================================================
Sub DelAllFavorite()
	If enchiasp.CheckPost=False Then
		ErrMsg = Postmsg
		Founderr = True
		Exit Sub
	End If
	enchiasp.Execute("Delete From ECCMS_Favorite where username='"& enchiasp.membername &"'")
	Call Returnsuc("<li>�ղؼ������ɣ�</li>")
End Sub
'================================================
' ��������SaveFavorite
' ��  �ã������ղ�
'================================================
Sub SaveFavorite()
	Call PreventRefresh
	If Trim(Request.Form("fondtopic")) = "" Then
		ErrMsg = ErrMsg + "<li>�ղصı��ⲻ��Ϊ�գ�</li>"
		Founderr = True
	End If
	If Trim(Request.Form("fondurl")) = "" Then
		ErrMsg = ErrMsg + "<li>�ղص�URL����Ϊ�գ�</li>"
		Founderr = True
	End If
	If CLng(GroupSetting(5)) <> 0 Then
		TotalNumber = enchiasp.Execute("Select Count(FavoriteID) from ECCMS_Favorite where username='"& enchiasp.CheckStr(enchiasp.membername) &"'")(0)
		If CLng(TotalNumber) >= CLng(GroupSetting(5)) Then
			ErrMsg = ErrMsg + "<li>�Բ��������ֻ���ղ�" & GroupSetting(5) & "����Ϣ��</li>"
			Founderr = True
		End If
	End  If
	If Founderr = True Then Exit Sub
	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL = "select * from [ECCMS_Favorite] where (FavoriteID is null)"
	Rs.Open SQL, Conn, 1, 3
	Rs.Addnew
		Rs("userid") = enchiasp.memberid
		Rs("username") = enchiasp.membername
		Rs("fondtopic") = Left(enchiasp.ChkFormStr(Request.Form("fondtopic")),80)
		Rs("fondurl") = Left(enchiasp.ChkFormStr(Request.Form("fondurl")),220)
		Rs("addTime") = Now()
	Rs.update
	Rs.Close:Set Rs = Nothing
	Call Returnsuc("<li>��ϲ��������ղسɹ���</li>")
End Sub
'================================================
' ��������AddFavorite
' ��  �ã�����ղ�
'================================================
Sub AddFavorite()
	Dim fondtopic,fondurl
	If Trim(Request("topic")) = "" Then
		ErrMsg = ErrMsg + "<li>�ղصı��ⲻ��Ϊ�գ�</li>"
		Founderr = True
	Else
		fondtopic = Trim(Request("topic"))
	End If
	If CLng(GroupSetting(5)) <> 0 Then
		TotalNumber = enchiasp.Execute("Select Count(FavoriteID) from ECCMS_Favorite where username='"& enchiasp.CheckStr(enchiasp.membername) &"'")(0)
		If CLng(TotalNumber) >= CLng(GroupSetting(5)) Then
			ErrMsg = ErrMsg + "<li>�Բ��������ֻ���ղ�" & GroupSetting(5) & "����Ϣ��</li>"
			Founderr = True
		End If
	End  If
	If Founderr = True Then Exit Sub
	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL = "select * from [ECCMS_Favorite] where (FavoriteID is null)"
	Rs.Open SQL, Conn, 1, 3
	Rs.Addnew
		Rs("userid") = enchiasp.memberid
		Rs("username") = enchiasp.membername
		Rs("fondtopic") = Left(enchiasp.ChkFormStr(Trim(fondtopic)),80)
		Rs("fondurl") = Left(Request.ServerVariables("HTTP_REFERER"),220)
		Rs("addTime") = Now()
	Rs.update
	Rs.Close:Set Rs = Nothing
	Call Returnsuc("<li>��ϲ��������ղسɹ���</li>")
End Sub
%>
<!--#include file="foot.inc"-->












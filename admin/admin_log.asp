<!--#include file =setup.asp-->
<!--#include file =check.asp-->
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
Dim LogType,WhereSQL,i
Dim TotalNumber,TotalPageNum,CurrentPage,maxperpage
Admin_header
If Not ChkAdmin("rizhi") Then
	Server.Transfer("showerr.asp")
	Response.End
End If
ConnectionLogDatabase
LogType = Trim(Request("type"))
Select Case Trim(LCase(Request("Action")))
	Case "del"
		Call BatchDel
	Case Else
		Call LogMain
End Select
If Founderr = True Then
	ReturnError(ErrMsg)
End If
Admin_footer
CloseConn
Sub Logmain()
%>
<TABLE width='98%' align=center cellpadding=3 cellspacing=1 border=0 class=tableBorder>
<TR>
	<TH colspan=5>��̨��־����</TH>
</TR>
<TR height=25>
	<TD colspan=5 class=TableRow2> <B>ѡ��鿴��־�¼���</B> 
	<A HREF=admin_log.asp>�鿴ȫ����־</A> | <A HREF=admin_log.asp?type=0>�鿴�¼�0</A> | <A HREF=admin_log.asp?type=1>�鿴�¼�1</A></TD>        
</TR>
<TR>
	<TH noWrap>����</TH>
	<TH noWrap>�� �� ��</TH>
	<TH noWrap> �� �� </TH>
	<TH width='70%'>�¼�����</TH>
	<TH noWrap>����ʱ��/IP</TH>
</TR><form action=admin_log.asp?action=del&LogType= method=post name=even>
<%
	Set Rs = Server.CreateObject("ADODB.Recordset")
	maxperpage = 20 '###ÿҳ��ʾ��
	If Not IsNumeric(Request("page")) And Len(Request("page")) <> 0 Then
		Response.Write ("�����ϵͳ����!����������")
		Response.Write
	End If
	If Not IsEmpty(Request("page")) And Len(Request("page")) <> 0 Then
		CurrentPage = CInt(Request("page"))
	Else
		CurrentPage = 1
	End If
	If LogType = "" Then
		WhereSQL = ""
	Else
		WhereSQL = "where LogType=" & LogType
	End If
	TotalNumber = lConn.Execute("Select count(logid) from [ECCMS_LogInfo] "& WhereSQL &"")(0)
	TotalPageNum = CInt(TotalNumber / maxperpage)  '�õ���ҳ��
	If TotalPageNum < TotalNumber / maxperpage Then TotalPageNum = TotalPageNum + 1
	If CurrentPage < 1 Then CurrentPage = 1
	If CurrentPage > TotalPageNum Then CurrentPage = TotalPageNum
	SQL = "select * from ECCMS_LogInfo "& WhereSQL &" order by logid desc"
	Rs.Open SQL, lConn, 1, 1
	If Rs.bof And Rs.EOF Then
		Response.Write "<tr><td colspan=5 class=TableRow1>��û���ҵ��κ���־��</td></tr>"
	Else
		If TotalPageNum > 1 then Rs.Move (CurrentPage - 1) * maxperpage
		i = 0
		Do While Not Rs.EOF And i < CInt(maxperpage)
%>
<TR height=22>
	<TD class=TableRow1 align=center><input type=checkbox name=logid value="<%=Rs("logid")%>"></TD>
	<TD class=TableRow1 noWrap><%=Server.HTMLEncode(Rs("username"))%></TD>
	<TD class=TableRow1 noWrap><%=Server.HTMLEncode(Rs("ScriptName"))%></TD>
	<TD class=TableRow1><%=Server.HTMLEncode(Rs("ActContent"))%></TD>
	<TD class=TableRow1 noWrap><%=Rs("LogAddTime")%><BR><%=Rs("UserIP")%></TD>
</TR>
<%
			Rs.movenext
			i = i + 1
			If i >= maxperpage Then Exit Do
		Loop
	End If
	Rs.Close:Set Rs = Nothing
%>
<TR height=25>
	<TD class=TableRow1 align=center colspan=5><B>��ѡ��Ҫɾ������־�¼�:<B> 
	<input type=checkbox name=chkall value=on onclick="CheckAll(this.form)"> ȫѡ         
	 <input class=Button type=submit name=act value=ɾ����־  onclick="{if(confirm('��ȷ��ִ�еĲ�����?')){this.document.even.submit();return true;}return false;}">
	 <input class=Button type=submit name=act onclick="{if(confirm('ȷ��������е���־��¼��?')){this.document.even.submit();return true;}return false;}" value=�����־>        
	 <input class=Button type=submit name=act onclick="{if(confirm('ȷ��ѹ����־���ݿ���?')){this.document.even.submit();return true;}return false;}" value=ѹ�����ݿ�></TD>        
</TR></form>
<TR height=25>
	<TD class=TableRow2 align=center colspan=5><%Call showpage%></TD>
</TR>
</TABLE>
<%
End Sub
Private Sub showpage()
	Dim n,ii
		If totalnumber Mod maxperpage = 0 Then
			n = totalnumber \ maxperpage
		Else
			n = totalnumber \ maxperpage + 1
		End If
		Response.Write "<table cellspacing=1 width='100%' border=0><form method=Post action=?type=" & Request("type") & "><tr><td align=center> " & vbCrLf
		If CurrentPage < 2 Then
			Response.Write "������־ <font COLOR=#FF0000><strong>" & totalnumber & "</strong></font> ��&nbsp;�� ҳ&nbsp;��һҳ&nbsp;"
		Else
			Response.Write "������־ <font COLOR=#FF0000><strong>" & totalnumber & "</strong></font> ��&nbsp;<a href=?page=1&type=" & Request("calsid") & ">�� ҳ</a>&nbsp;"
			Response.Write "<a href=?page=" & CurrentPage - 1 & "&type=" & Request("type") & ">��һҳ</a>&nbsp;"
		End If
		If n - CurrentPage < 1 Then
			Response.Write "��һҳ&nbsp;β ҳ" & vbCrLf
		Else
			Response.Write "<a href=?page=" & (CurrentPage + 1) & "&type=" & Request("type") & ">��һҳ</a>"
			Response.Write "&nbsp;<a href=?page=" & n & "&type=" & Request("type") & ">β ҳ</a>" & vbCrLf
		End If
		Response.Write "&nbsp;ҳ�Σ�<strong><font color=red>" & CurrentPage & "</font>/" & n & "</strong>ҳ "
		Response.Write "&nbsp;ת����"
			Response.Write "&nbsp;<select name='page' size='1' style=""font-size: 9pt"" onChange='javascript:submit()'>" & vbCrLf
			For ii = 1 To n
				Response.Write "<option value='" & ii & "' "
				If CurrentPage = CInt(ii) Then
					Response.Write "selected "
				End If
				Response.Write ">��" & ii & "ҳ</option>"
			Next
			Response.Write "</select> " & vbCrLf
		Response.Write "</td></tr></FORM></table>" & vbCrLf
End Sub

Sub BatchDel()
	Dim logid
	If Request("act")="ɾ����־" Then
		If request.form("logid")="" Then
			ErrMsg =  "��ָ������¼���"
			Founderr = True
			Exit Sub
		Else
			logid=replace(Request.Form("logid"),"'","")
			logid=replace(logid,";","")
			logid=replace(logid,"--","")
			logid=replace(logid,")","")
		End If
	End If
	If Request("act")="ѹ�����ݿ�" Then
		If CompressMDB("Logdata.Asa") Then OutHintScript("��ϲ�� ^_^ ��־���ݿ�ѹ���ɹ���")
	ElseIf Request("act")="ɾ����־" Then
			lConn.Execute("delete from ECCMS_LogInfo where Datediff('D',LogAddTime, Now()) > 3 And logid in ("&logid&")")
	ElseIf Request("act")="�����־" Then
		If request("LogType")="" or IsNull(Request("LogType")) Then 
			lConn.Execute("delete from ECCMS_LogInfo Where Datediff('D',LogAddTime, Now) > 3")
		Else
			lConn.Execute("delete from ECCMS_LogInfo where  Datediff('D',LogAddTime, Now()) > 3 And LogType="&CInt(request("LogType"))&"")
		End If
	End If
	Succeed ("�ɹ�ɾ����־��ע�⣺�����ڵ���־�ᱻϵͳ������")
End Sub
'================================================
' ��������CompressMDB
' ��  �ã�ѹ��ACCESS���ݿ�
' ��  ����dbPath ----���ݿ�·��
' ����ֵ��True  ----  False
'================================================
Public Function CompressMDB(DBPath)
        Dim fso, Engine, strDBPath, JET_3X
        CompressMDB = False
        If DBPath = "" Then Exit Function
        If InStr(DBPath, ":") = 0 Then DBPath = Server.MapPath(DBPath)
        strDBPath = Left(DBPath, InStrRev(DBPath, "\"))
        Set fso = CreateObject(enchiasp.FSO_ScriptName)

        If fso.FileExists(DBPath) Then
                fso.CopyFile DBPath, strDBPath & "temp.mdb"
                Set Engine = CreateObject("JRO.JetEngine")

                Engine.CompactDatabase "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBPath & "temp.mdb", _
                "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBPath & "temp1.mdb"

                fso.CopyFile strDBPath & "temp1.mdb", DBPath
                fso.DeleteFile (strDBPath & "temp.mdb")
                fso.DeleteFile (strDBPath & "temp1.mdb")
                Set fso = Nothing
                Set Engine = Nothing
                CompressMDB = True
        Else
                CompressMDB = False
        End If
End Function
%>
<script language="javascript">
function CheckAll(form)  
  {  
  for (var i=0;i<form.elements.length;i++)  
    {  
    var e = form.elements[i];  
    if (e.name != 'chkall')  
       e.checked = form.chkall.checked;  
    }  
  }  
</script>
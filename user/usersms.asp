<!--#include file="config.asp"-->
<!--#include file="check.asp"-->
<!--#include file="head.inc"-->
<%
'=====================================================================
' ������ƣ�������վ����ϵͳ---�û����ŷ���
' ��ǰ�汾��enchicms Version 3.0.0
' �������ڣ�2005-03-25
' �ٷ���վ���˳��ж�������Ƽ��������޹�˾(www.enchi.com.cn) 
' ����֧�֣����Ʒ�
' ���䣺liuyunfan@163.com
' QQ��21556923
'=====================================================================
' Copyright 2005-2008  All Rights Reserved.
'=====================================================================
Call InnerLocation("�û����ŷ���")

Dim Rs,SQL,i,Action
Dim Maxsms,boxname,smstype,readaction

If CInt(GroupSetting(22)) = 0 Then
	ErrMsg = ErrMsg + "<li>�Բ�����û��ʹ�ö��ŷ����Ȩ�ޣ�����ʲô��������ϵ����Ա��</li>"
	Founderr = True
End If
Maxsms = CLng(GroupSetting(24))
Call showmain
If Founderr = True Then
	Call Returnerr(ErrMsg)
End If

Sub showmain()
	If Founderr = True Then Exit Sub
	Dim smsCount,DelCount
	smsCount=0
	Set Rs = enchiasp.Execute("select Count(id) from ECCMS_Message Where flag=0 And incept='"& enchiasp.membername &"'")
	smsCount = CLng(Rs(0))
	'�����ж�Ϊ�Զ�ɾ��������Ķ���Ϣ
	If smsCount > Maxsms And Maxsms <> 0 Then
		i = smsCount-Maxsms
		Set Rs=enchiasp.Execute("select top "& i &" id from ECCMS_Message Where incept='"& enchiasp.membername &"' Order by id,isRead Desc")
		While Not Rs.EOF
			enchiasp.Execute("Delete from ECCMS_Message Where id="& rs(0))
			Rs.movenext
		Wend
		smsCount = Maxsms
	End if
	Rs.Close:Set Rs = Nothing
%>
<script language="JavaScript">
<!--
function enchiasp_usersms_smsbox_top(smstype){
	document.write ('<th valign=middle width=30 height=25 noWrap>�Ѷ�</th>');
	document.write ('<th valign=middle width=100>');
	if (smstype=='inbox')
	{
		document.write ('������');
	}else{
		document.write ('�ռ���');
	}
	document.write ('</th>');
	document.write ('<th valign=middle width=300>����</th>');
	document.write ('<th valign=middle width=150>����</th>');
	document.write ('<th valign=middle width=50>��С</th>');
	document.write ('<th valign=middle width=30 noWrap>����</th>');
	document.write ('</tr>');
}
function enchiasp_usersms_smsbox_emp(boxname){
	document.write ('<tr>');
	document.write ('<td class=Usertablerow1 align=center valign=middle colspan=6>����'+boxname+'��û���κ����ݡ�</td>');
	document.write ('</tr>');
}
function enchiasp_usersms_smsbox_loop(flag,isread,sms_type,sender,incept,title,sendtime,clength,id,readaction){
	var tablebody,newstyle;
	if (isread==0)
	{
		tablebody="Usertablerow2";
		newstyle="font-weight:bold";
	}else{
		tablebody="Usertablerow1";
		newstyle="font-weight:normal";
	}
	document.write ('<tr>');
	document.write ('<td class='+tablebody+' align=center valign=middle>');
	if (flag==0){
		if (isread==0){
			document.write ('<img src="images/m_news.gif" border=0 alt="�¶���">');
			}else{
			document.write ('<img src="images/m_olds.gif" border=0 alt="�ɶ���">');
		}
	}else{
		document.write ('<img src="images/m_issend_2.gif" border=0 alt="ϵͳ����">');
	}
	document.write ('</td>');
	document.write ('<td noWrap class='+tablebody+' align=center valign=middle style="'+newstyle+'">');
	if (sms_type=='inbox')
	{
		document.write ('<a href="userlist.asp?name='+sender+'" target=_blank>'+sender+'</a>');
	}else
	{
		document.write ('<a href="userlist.asp?name='+incept+'" target=_blank>'+incept+'</a>');
	}
	document.write ('</td>');
	document.write ('<td class='+tablebody+' align=left style="'+newstyle+'"><a href="message.asp?action='+readaction+'&sid='+id+'&sender='+sender+'">'+title+'</a>	</td>');
	document.write ('<td noWrap class='+tablebody+' style="'+newstyle+'">'+sendtime+'</td>');
	document.write ('<td noWrap class='+tablebody+' style="'+newstyle+'">'+clength+'Byte</td>');
	document.write ('<td align=center valign=middle width=30 class='+tablebody+'><input type=checkbox name=id value='+id+'></td>');
	document.write ('</tr>');
}
function enchiasp_usersms_smsbox_footer(boxname){
	document.write ('<tr>');
	document.write ('<td align=right valign=middle colspan=6 class=Usertablerow2>��ʡÿһ�ֿռ䣬�뼰ʱɾ��������Ϣ&nbsp;<input type=checkbox name=chkall value=on onclick="CheckAll2(this.form)">ѡ��������ʾ��¼&nbsp;<input type=submit name=action onclick="{if(confirm(\'ȷ��ɾ��ѡ���ļ�¼��?\')){return true;}return false;}" value="ɾ��'+boxname+'" class=button>&nbsp;<input type=submit name=action onclick="{if(confirm(\'ȷ�����'+boxname+'���еļ�¼��?\')){this.document.inbox.submit();return true;}return false;}" value="���'+boxname+'" class=button></td>');
	document.write ('</tr>');
	document.write ('</form></table>');
}
//-->
</script>
<table cellspacing=1 align=center cellpadding=3 border=0 class=Usertableborder>
	<tr>
		<th>>> ���ŷ��� <<</th>
	</tr>
	<tr>
		<td align=center class=Usertablerow1><a href="usersms.asp?action=inbox"><img src="images/m_inbox.gif" border="0" alt="�ռ���"></a>&nbsp;
		<a href="usersms.asp?action=sendbox"><img src="images/M_issend.gif" border="0" alt="�ѷ����ʼ�"></a>&nbsp;
		<a href="message.asp?action=alldel" onclick=showClick('��ȷ��Ҫ������ж���Ϣ��?')><img src="images/recycle.gif" border="0" alt="������ж���Ϣ"></a>&nbsp;
		<a href="friend.asp"><img src="images/M_address.gif" border="0" alt="��ַ��"></a>&nbsp;
		<a href="message.asp?action=new"><img src="images/m_write.gif" border="0" alt="����ѶϢ"></a></td>
	</tr>
</table>
<br style="overflow: hidden; line-height: 10px">
<table cellspacing=1 align=center cellpadding=3 bgcolor=#cccccc border=0 class=Usertableborder>
	<tr height=20>
		<td colspan=6 class=Usertablerow1><table Width="100%" cellpadding=2 cellspacing=1 border=0 align=center style="display:nowrap"><TR>
<td Width="100" align=right>��������������</td>
<td Width="*"><img src="images/bar1.gif" width="0" height="16" id="Sms_bar" align=absmiddle></td>
<td Width="150" align=center id="Sms_txt">0%</td>
</tr></table></td>
	</tr>
	<form action="message.asp" method=post name=inbox>
<%
	SQL = "select * from ECCMS_Message "
	Action = LCase(Request("action"))
	Select Case Trim(Action)
		Case "inbox"
			SQL = SQL + " where incept = '"& enchiasp.membername &"' Or flag = 1 order by id desc"
			boxname = "�ռ���"
			smstype = "inbox"
			readaction = "read"
		Case "sendbox"
			SQL = SQL + " where sender = '"& enchiasp.membername &"' And delSend = 0 order by id desc"
			boxname = "������"
			smstype = "sendbox"
			readaction = "outread"
		Case Else
			SQL = SQL + " where incept = '"& enchiasp.membername &"' Or flag = 1 order by id desc"
			boxname = "�ռ���"
			smstype = "inbox"
			readaction = "read"
	End Select
	Call usersmsbox
	Response.Write ShowTable("Sms_bar","Sms_txt",smsCount,Maxsms)
End Sub
'================================================
' ��������usersmsbox
' ��  �ã��û������б�
'================================================
Sub usersmsbox()
	Dim newstyle
	Dim CurrentPage,page_count,totalrec,Pcount,PageListNum
	PageListNum = 20
	CurrentPage = Request("page")
	If CurrentPage = "" Or Not IsNumeric(CurrentPage) Then
		CurrentPage = 1
	Else
		CurrentPage = CLng(CurrentPage)
	End If
	Response.Write "<script>enchiasp_usersms_smsbox_top('"& smstype &"')</script>"
	Set Rs = Server.CreateObject("ADODB.Recordset")
	Rs.Open SQL,conn,1,1
	If Rs.EOF And Rs.BOF Then
		Response.Write "<script>enchiasp_usersms_smsbox_emp('"& boxname &"')</script>"
	Else
		Rs.PageSize = PageListNum
		Rs.AbsolutePage = CurrentPage
		page_count = 0
		totalrec = Rs.recordcount
		Do While Not Rs.EOF And (Not page_count = Rs.PageSize)
			Response.Write VbCrLf
			Response.Write "<script>enchiasp_usersms_smsbox_loop("
			Response.Write Rs("flag")
			Response.Write ","
			Response.Write Rs("isRead")
			Response.Write ",'"
			Response.Write smstype
			Response.Write "','"
			Response.Write EncodeJS(Rs("sender"))
			Response.Write "','"
			Response.Write EncodeJS(Rs("incept"))
			Response.Write "','"
			Response.Write EncodeJS(Rs("title"))
			Response.Write "','"
			Response.Write Rs("sendtime")
			Response.Write "',"
			Response.Write Len(Rs("content"))
			Response.Write ","
			Response.Write Rs("id")
			Response.Write ",'"
			Response.Write readaction
			Response.Write "')</script>"
			Response.Write VbCrLf
			page_count = page_count + 1
		Rs.movenext
		Loop
	End If
	Rs.close:Set Rs = nothing
	If totalrec Mod PageListNum = 0 Then
		Pcount =  totalrec \ PageListNum
	Else
		Pcount =  totalrec \ PageListNum+1
	End If
	If page_count = 0 Then CurrentPage = 0
	Response.Write "	<tr height=20>" & vbNewLine
	Response.Write "		<td colspan=6 class=Usertablerow1>"
	Response.Write ShowPages(CurrentPage,Pcount,totalrec,PageListNum,"action="& Request("action"))
	Response.Write "</td>"
	Response.Write "	</tr>" & vbNewLine
	Response.Write VbCrLf
	Response.Write "<script>enchiasp_usersms_smsbox_footer('"& boxname &"')</script>"
End Sub
'================================================
' ��������ShowTable
' ��  �ã���ʾ��������
' ��  ������ͼƬ�������ƣ�����������ƣ���������������
'================================================
Function ShowTable(SrcName,TxtName,str,c)
	Dim Tempstr,Src_js,Txt_js,TempPercent
	Tempstr = str/C
	TempPercent = FormatPercent(tempstr,0,-1)
	Src_js = "document.getElementById(""" + SrcName + """)"
	Txt_js = "document.getElementById(""" + TxtName + """)"
	ShowTable = VbCrLf + "<script>"
	ShowTable = ShowTable + Src_js + ".width=""" & FormatNumber(tempstr*300,0,-1) & """;"
	ShowTable = ShowTable + Src_js + ".title=""��������Ϊ��"&c&"�����ܹ��Ѵ��棨"&str&"�������ţ�"";"
	ShowTable = ShowTable + Txt_js + ".innerHTML="""
	If FormatNumber(tempstr*100,0,-1) < 80 Then
		ShowTable = ShowTable + "��ʹ��:" & TempPercent & """;"
	Else
		ShowTable = ShowTable + "<font color=\""red\"">��ʹ��:" & TempPercent & ",��Ͽ�����</font>"";"
	End If
	ShowTable = ShowTable + "</script>"
End Function
Function EncodeJS(str)
	str = enchiasp.HtmlEncode(str)
	str = Replace(Replace(Replace(Replace(str,"\","\\"),"'","\'"),VbCrLf,"\n"),Chr(13),"")
	EnCodeJs = str
End Function
%><!--#include file="foot.inc"-->






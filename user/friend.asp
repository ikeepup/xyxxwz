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
Call InnerLocation("���ѹ���")

Dim Rs,SQL,i

If CInt(GroupSetting(4)) = 0 Then
	ErrMsg = ErrMsg + "<li>�Բ�����û��ʹ�ú��ѹ����Ȩ�ޣ�����ʲô��������ϵ����Ա��</li>"
	Founderr = True
End If
Dim Action:Action = LCase(Request("action"))
Select Case Trim(Action)
	Case "add"
		Call FriendAdd
	Case "�ƶ�"
		Call MoveFriend
	Case "ɾ��"
		Call FriendDel
	Case "��պ���"
		Call DelAllFriend
	Case Else
		Call showmain
End Select
If Founderr = True Then
	Call Returnerr(ErrMsg)
End If
Sub showmain()
	If Founderr = True Then Exit Sub
	Dim PageListNum,totalrec,Pcount,CurrentPage,page_count
	PageListNum = 20
	page_count = 0
	If Not IsNumeric(Request("page")) And Trim(Request("page")) <> "" Then
		Response.Write ("�����ϵͳ����!����������")
		Response.End
	End If
	If Not IsEmpty(Request("page")) And Trim(Request("page")) <> "" Then
		CurrentPage = Clng(Request("page"))
	Else
		CurrentPage = 1
	End If
	totalrec = enchiasp.Execute("Select Count(FriendID) from ECCMS_Friend where username='"& enchiasp.CheckStr(enchiasp.membername) &"'")(0)
	If totalrec Mod PageListNum = 0 Then
		Pcount = totalrec \ PageListNum
	Else
		Pcount = totalrec \ PageListNum + 1
	End If
	If CurrentPage > Pcount Then CurrentPage = Pcount
	If CurrentPage < 1 Then CurrentPage = 1
%>
<table cellspacing=1 align=center cellpadding=3 border=0 class=Usertableborder>
	<tr>
		<th colspan=6>>> �ҵĺ��� <<</th>
	</tr>
	<form action="friend.asp" method=post name=inbox>
	<tr>
		<td width="15%" align=center class=Usertablerow2><b class=userfont2>�� ��</b></td>
		<td width="25%" align=center class=Usertablerow2><b class=userfont2>�û���</b></td>
		<td width="30%" align=center class=Usertablerow2><b class=userfont2>�� ��</b></td>
		<td width="10%" align=center class=Usertablerow2><b class=userfont2>OICQ</b></td>
		<td width="10%" align=center class=Usertablerow2><b class=userfont2>������</b></td>
		<td width="10%" align=center class=Usertablerow2><b class=userfont2>�� ��</b></td>
	</tr>
<%
	Set Rs=Server.Createobject("adodb.recordset")
	SQL = "select F.FriendID,F.userid,F.Friend,F.grouping,U.usermail,U.HomePage,U.oicq From [ECCMS_Friend] F inner join [ECCMS_User] U on F.Friend=U.username where F.userid="&enchiasp.memberid
	SQL = SQL+" order by F.addtime desc"
	Rs.Open SQL,Conn,1,1
	If Rs.EOF And Rs.BOF Then
		Rs.Close:Set Rs = Nothing
	Else
		Rs.Move (CurrentPage - 1) * Cint(PageListNum)
		SQL = Rs.GetRows(PageListNum)
		Rs.Close:Set Rs = Nothing
		For i=0 To Ubound(SQL,2)
%>
	<tr>
		<td align=center class=Usertablerow1><b class=userfont2><%
		If CInt(SQL(3,i)) = 0 Then
			Response.Write "İ����"
		ElseIf CInt(SQL(3,i)) = 1 Then
			Response.Write "�ҵĺ���"
		ElseIf CInt(SQL(3,i)) = 2 Then
			Response.Write "������"
		Else
			Response.Write "������"
		End If
		%></b></td>
		<td align=center class=Usertablerow1><a href="dispuser.asp?name=<%=SQL(2,i)%>" target=_blank title="��� <%=SQL(2,i)%> �ĸ�������"><%=SQL(2,i)%></a></td>
		<td align=center class=Usertablerow1><a href="mailto:<%=SQL(4,i)%>"><%=SQL(4,i)%></a></td>
		<td align=center class=Usertablerow1><a href="http://search.tencent.com/cgi-bin/friend/user_show_info?ln=<%=SQL(6,i)%>" title="<%=SQL(2,i)%> �� Oicq:<%=SQL(6,i)%>" target=_blank><img src=images/oicq.gif border=0></a></td>
		<td align=center class=Usertablerow1><a href="message.asp?action=new&touser=<%=SQL(2,i)%>" title="�� <%=SQL(2,i)%> ������"><img src=images/message.gif border=0></a></td>
		<td align=center class=Usertablerow1><input type=checkbox name=id value="<%=SQL(0,i)%>"></td>
	</tr>
<%
			page_count = page_count+1
		Next
	End If
%>
	<tr>
		<td colspan=6 align=center class=Usertablerow1><%Response.Write ShowPages (CurrentPage,Pcount,totalrec,PageListNum,"")%></td>
	</tr>
	<tr>
		<td colspan=6 align=center class=Usertablerow2><input type=checkbox name=chkall value=on onclick="CheckAll2(this.form)">ѡ��������ʾ��¼&nbsp;
		<select name="grouping">
		<option value="" selected>�����ƶ���...</option>
		<option value="0" >İ����</option>
		<option value="1" >�ҵĺ���</option>
		<option value="2" >������</option>
		</select>&nbsp;
		<input type=submit name=action onclick="{if(confirm('ȷ���ƶ�ѡ���ļ�¼��?')){this.document.inbox.submit();return true;}return false;}" value="�ƶ�" class=button>&nbsp;
		<input type=button name=action onclick="showsub('addfriend')" value="��Ӻ���" class=button>&nbsp;
		<input type=submit name=action onclick="{if(confirm('ȷ��ɾ��ѡ���ļ�¼��?')){this.document.inbox.submit();return true;}return false;}" value="ɾ��" class=button>&nbsp;
		<input type=submit name=action onclick="{if(confirm('ȷ��������еļ�¼��?')){this.document.inbox.submit();return true;}return false;}" value="��պ���" class=button></td>
	</tr></form>
</table>
<div id=addfriend style="display:none">
<br style="overflow: hidden; line-height: 10px">
<table cellspacing=1 align=center cellpadding=3 border=0 class=Usertableborder>
	<tr>
		<th>>> ��Ӻ��� <<</th>
	</tr>
	<form name=myform method=post action=?action=add>
	<tr>
		<td align=center class=Usertablerow1><b class=userfont2>���ѣ�</b><input type="text" name="friend" size=45>
		<b class=userfont2>���</b><select name="grouping">
		<option value="0" selected>��ѡ��....</option>
		<option value="0" >İ����</option>
		<option value="1" >�ҵĺ���</option>
		<option value="2" >������</option>
		</select>
		<input type=submit value="���" class=button>&nbsp;<input type="reset" name="Clear" value="���" class=button><br>
		<div><b>ע�⣺</b><%If CLng(GroupSetting(6)) <> 0 Then%>�����ֻ����� <b class=userfont1><%=GroupSetting(6)%></b> λ���ѣ�<%End If%>�������飬�����������Ժ������Ķ��š� </div></td>
	</tr>
	</form>
</table>
</div>

<%
End Sub
'================================================
' ��������FriendDel
' ��  �ã�����ɾ������
'================================================
Sub FriendDel()
	If enchiasp.CheckPost=False Then
		ErrMsg = Postmsg
		Founderr = True
		Exit Sub
	End If
	Dim FriendID,fixid
	FriendID = Replace(Request.form("id"),"'","")
	FriendID = Replace(FriendID,";","")
	FriendID = Replace(FriendID,"--","")
	FriendID = Replace(FriendID,")","")
	fixid = Replace(FriendID,",","")
	fixid = Trim(Replace(fixid," ",""))
	If FriendID = "" Or IsNull(FriendID) Then
		ErrMsg = ErrMsg + "<li>��Ч��ϵͳ������</li>"
		Founderr = True
		Exit Sub
	ElseIf Not IsNumeric(fixid) Then
		ErrMsg = ErrMsg + "<li>��Ч��ϵͳ������</li>"
		Founderr = True
		Exit Sub
	Else
		enchiasp.Execute("Delete From ECCMS_Friend where userid="&enchiasp.memberid&" And FriendID in ("&FriendID&")")
		Call Returnsuc("<li>����ɾ���ɹ���</li>")
	End If
End Sub
'================================================
' ��������DelAllFriend
' ��  �ã�ɾ�����к���
'================================================
Sub DelAllFriend()
	If enchiasp.CheckPost=False Then
		ErrMsg = Postmsg
		Founderr = True
		Exit Sub
	End If
	enchiasp.Execute("Delete From ECCMS_Friend where userid="& enchiasp.memberid)
	Call Returnsuc("<li>������ճɹ���</li>")
End Sub
'================================================
' ��������FriendAdd
' ��  �ã���Ӻ���
'================================================
Sub FriendAdd()
	Call PreventRefresh
	Dim grouping,strIncept,FriendName,TotalFriend
	If enchiasp.CheckPost=False Then
		ErrMsg = Postmsg
		Founderr = True
		Exit Sub
	End If
	If Trim(Request("friend")) = "" Then
		ErrMsg = ErrMsg + "<li>��ѡ��Ҫ��Ӻ��ѵ����ƣ�</li>"
		Founderr = True
	Else
		strIncept = enchiasp.CheckBadstr(Request("friend"))
		strIncept = split(strIncept,",")
	End If
	If Trim(Request("grouping"))<>"" And IsNumeric(Request("grouping")) then 
		grouping = CInt(Request("grouping"))
	Else
		grouping = 0
	End If
	If Founderr = True Then Exit Sub
	For i = 0 To Ubound(strIncept)
		If i >= 5 Then Exit For
		FriendName = Trim(strIncept(i))
		SQL="select username from [ECCMS_User] where username='"&FriendName&"'"
		Set Rs = enchiasp.Execute(SQL)
		If Rs.EOF And Rs.BOF Then
			ErrMsg = ErrMsg + "<li>û���ҵ�<font color=red>" & FriendName & "</font>����û�������δ�ɹ���</li>"
			Founderr = True
			Exit Sub
		Else
			FriendName = Rs(0)
		End If
		Rs.close
		If enchiasp.membername = Trim(FriendName) Then
			ErrMsg = ErrMsg + "<li>�Բ��𣡲��ܰ��������Ϊ���ѡ�</li>"
			Founderr = True
			Exit Sub
		End If
		If CLng(GroupSetting(6)) <> 0 Then
			TotalFriend = enchiasp.Execute("Select Count(FriendID) from ECCMS_Friend where userid="& enchiasp.memberid &" And username='"& enchiasp.CheckStr(enchiasp.membername) &"'")(0)
			If CLng(TotalFriend) >= CLng(GroupSetting(6)) Then
				ErrMsg = ErrMsg + "<li>�Բ��������ֻ����� <font color=red><b>" & GroupSetting(6) & "</b></font> λ���ѡ�</li>"
				Founderr = True
				Exit Sub
			End If
		End  If
		SQL = "Select FriendID From ECCMS_Friend Where userid="& enchiasp.memberid &" And friend='"& FriendName &"'"
		Set Rs = enchiasp.Execute(SQL)
		If Rs.EOF And Rs.BOF Then
			SQL = "Insert into ECCMS_Friend (userid,UserName,Friend,addTime,grouping) values ("& enchiasp.memberid &",'"& enchiasp.membername &"','"& FriendName &"',"& NowString &","& grouping &") "
			enchiasp.Execute(SQL)
		Else
			ErrMsg = ErrMsg + "<li><font color=red>" & FriendName & "</font>����û��Ѿ���ӹ��ˣ��벻Ҫ�ظ���ӣ�лл����</li>"
			Founderr = True
			Exit Sub
		End If
	Next
	Call Returnsuc("<li>��ϲ������Ӻ��ѳɹ���</li>")
End Sub
'================================================
' ��������MoveFriend
' ��  �ã��ƶ����ѵ�������
'================================================
Sub MoveFriend()
	If enchiasp.CheckPost=False Then
		ErrMsg = Postmsg
		Founderr = True
		Exit Sub
	End If
	Dim grouping
	Dim FriendID,fixid
	If Trim(Request("grouping"))<>"" And IsNumeric(Request("grouping")) Then
		grouping = CInt(Request("grouping"))
	Else
		ErrMsg = ErrMsg + "<li>���ѷ��鲻��Ϊ�ա�</li>"
		Founderr = True
		Exit Sub
	End If
	FriendID = Replace(Request.form("id"),"'","")
	FriendID = Replace(FriendID,";","")
	FriendID = Replace(FriendID,"--","")
	FriendID = Replace(FriendID,")","")
	fixid = Replace(FriendID,",","")
	fixid = Trim(Replace(fixid," ",""))
	If FriendID = "" Or IsNull(FriendID) Then
		ErrMsg = ErrMsg + "<li>��Ч��ϵͳ������</li>"
		Founderr = True
		Exit Sub
	ElseIf Not IsNumeric(fixid) Then
		ErrMsg = ErrMsg + "<li>��Ч��ϵͳ������</li>"
		Founderr = True
		Exit Sub
	Else
		enchiasp.Execute("Update ECCMS_Friend set grouping = "&grouping&" where userid="&enchiasp.memberid&" And FriendID in ("&FriendID&")")
		Call Returnsuc("<li>��ϲ�����ƶ����ѷ���ɹ���</li>")
	End If
End Sub
%>
<!--#include file="foot.inc"-->
















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
Dim GroupSetting,Action,i,strClass
Action = LCase(Request("action"))
If Not ChkAdmin("UserGroup") Then
	Server.Transfer("showerr.asp")
	Response.End
End If
select case Action
case "save" 
	call savegroup()
case "savedit" 
	call savedit()
case "del"
	call delgroup()
case "group" 
	call gradeinfo()
case "addgroup" 
	call addgroup()
case "editgroup"
	call editgroup()
case else
	call usergroup()
end select
If FoundErr = True Then
	ReturnError(ErrMsg)
End If
Admin_footer
SaveLogInfo(AdminName)
CloseConn
sub usergroup()
%>
<table width="98%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
<tr> 
<th height="23" colspan="4" >�û������&nbsp;&nbsp;|&nbsp;&nbsp;<a href="?action=addgroup"><font color=#FFFFFF><strong>[������û���]</strong></font></a></th>
</tr>
<tr><td colspan=4 height=25 class="tablerow1"><B>˵��</B>��<BR>
�����������������ø����û�����ϵͳ�е�Ĭ��Ȩ�ޣ�ϵͳĬ���û��鲻��ɾ���ͱ༭�û��ȼ�;<BR>
�ڿ��Խ�������û��������������Ȩ�ޣ����Խ��������û�ת�Ƶ����飬�뵽�û������н�����ز���;<BR>
�ۿ���ɾ���ͱ༭����ӵ��û��飬�����ʱ������Ӧ�û��ȼ���<BR>
</td></tr>
<tr align=center>
<td height="23" width="30%" class=TableTitle><B>�û���</B></td>
<td height="23" width="20%" class=TableTitle><B>�û�����</B></td>
<td height="23" width="20%" class=TableTitle><B>�༭</B></td>
<td height="23" width="30%" class=TableTitle><B>�û��ȼ�</B></td>
</tr>
<%
dim trs
set rs=enchiasp.execute("select * from ECCMS_UserGroup order by groupid")
i = 0
do while not rs.eof
set trs=enchiasp.execute("select count(userid) from [ECCMS_User] where UserGrade="&rs("Grades"))
	If (i mod 2) = 0 Then
		strClass = "class=TableRow1"
	Else
		strClass = "class=TableRow2"
	End If
%>
<tr align=center>
<td height="23" <%=strClass%>><%=rs("GroupName")%></td>
<td height="23" <%=strClass%>><%if rs("Grades") = 0 Then%>�����û�<%Else%><%=trs(0)%><%End If%></td>
<td height="23" <%=strClass%>><a href="?action=editgroup&groupid=<%=rs("groupid")%>">�û�������</a><%if rs("groupid") => 6 then%> | <a href="?action=del&groupid=<%=rs("groupid")%>&Grade=<%=rs("Grades")%>" onclick="{if(confirm('�˲�����ɾ�����û���\n ��ȷ��ִ�еĲ�����?')){return true;}return false;}">ɾ��</a><%end if%></td>
<td height="23" <%=strClass%>><%=rs("Grades")%></td>
</tr>
<%
rs.movenext
i = i + 1
loop
rs.close
set rs=nothing
%>
</table><BR>
<%
end sub
Sub addgroup()
        Dim GroupNum
        Set Rs = CreateObject("Adodb.recordset")
        SQL = "select Max(groupid) from ECCMS_UserGroup"
        Rs.Open SQL, Conn, 1, 1
        If Rs.EOF And Rs.bof Then
                GroupNum = 1
        Else
                GroupNum = Rs(0) + 1
        End If
        If IsNull(GroupNum) Then GroupNum = 1
        Rs.Close
%>
<table width="98%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
<tr> 
<th height="23" colspan="2" >����µ��û���</th>
</tr>
<FORM METHOD=POST ACTION="admin_group.asp?action=save">
<input type="hidden" name="newgroupid" value="<% = GroupNum %>">
<tr><td colspan=2 height=25 class="tablerow1"><B>˵��</B>��<BR>
�ٿ��Խ�������û��������������Ȩ�ޣ����Խ��������û�ת�Ƶ����飬�뵽�û������н�����ز�����<BR>
�ڿ���ɾ���ͱ༭����ӵ��û��飬���������д��Ӧ�û��ȼ���<BR>
</td></tr>
<tr> 
<th colspan="2" >����µ��û���</th>
</tr>
<tr>
<td width="60%" class=tablerow1>�û�������</td>
<td width="40%" class=tablerow1><input size=35 name="GroupName" type=text></td>
</tr>
<script>function checkclick(msg){if(confirm(msg)){event.returnValue=true;}else{event.returnValue=false;}}</script>
<tr>
<td class=tablerow1>�û���ȼ�������������(����Խ�󼶱�Խ��)</td>
<td class=tablerow1><input size=10 name="Grades" type=text value=<%=conn.execute("Select max(Grades)from ECCMS_UserGroup where Grades <> 999")(0)+1%>></td>
</tr>
<tr> 
<td class=tablerow1>
</td>
<td class=tablerow1>
<input type="button" name="Submit1" onclick="javascript:history.go(-1)" value="������һҳ" class=button>��
<input type="submit" name="submit" value="����û���" class=button></td>
</tr>
</FORM>
</table><BR>
<%
set rs=nothing
End Sub
Sub editgroup()
	Dim GroupSet
	SQL = "select groupid,GroupName,GroupSet,Grades from ECCMS_UserGroup where groupid = " & Request("groupid")
	Set Rs = enchiasp.Execute(SQL)
	GroupSet = Split(Rs("GroupSet"),"|||")
%>
<table border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
  <tr> 
    <th colspan="2" >�޸��û���</th>
  </tr>
  <form method=post action="admin_group.asp?action=savedit">
  <tr>
    <td colspan=2 height=25 class="tablerow1"><B>˵��</B>��<BR>
    �ٿ��Խ����޸��û��������������Ȩ�ޣ����Խ��������û�ת�Ƶ����飬�뵽�û������н�����ز�����<BR>
    </td>
  </tr>
  <tr> 
    <th colspan="2">�û�������</th>
  </tr>
  <tr>
    <td width="60%" class=tablerow1>�û�������</td>
    <td width="40%" class=tablerow1><input size=35 name="GroupName" type=text value="<%=Rs("GroupName")%>"></td>
  </tr>
  <tr>
    <td class=tablerow2>�û���ȼ�������������(<font color=blue>����Խ�󼶱�Խ��</font>)</td>
    <td class=tablerow2><input size=10 type=text value="<%=Rs("Grades")%>" disabled>
    <input size=10 name="Grades" type=hidden value="<%=Rs("Grades")%>">&nbsp;&nbsp;&nbsp;&nbsp;
    <a href="admin_group.asp">�����û�����ҳ</a></td>
  </tr>
  <tr> 
    <th colspan="2" align=left>�����û�����ʹ������</th>
  </tr>
  <tr>
    <td class=tablerow1>�û��Ƿ�����޸�����</td>
    <td class=tablerow1><input type=radio name="GroupSet(0)" value=0<%If CInt(GroupSet(0)) = 0 Then Response.Write " checked"%>> ��&nbsp;&nbsp;
      <input type=radio name="GroupSet(0)" value=1<%If CInt(GroupSet(0)) = 1 Then Response.Write " checked"%>> �� </td>
  </tr>
  <tr>
    <td class=tablerow2>�û��Ƿ�����޸�����</td>
    <td class=tablerow2><input type=radio name="GroupSet(1)" value=0<%If CInt(GroupSet(1)) = 0 Then Response.Write " checked"%>> ��&nbsp;&nbsp;
      <input type=radio name="GroupSet(1)" value=1<%If CInt(GroupSet(1)) = 1 Then Response.Write " checked"%>> �� </td>
  </tr>
  <tr>
    <td class=tablerow1>����������Ϣ�Ƿ�ʹ����֤��</td>
    <td class=tablerow1><input type=radio name="GroupSet(2)" value=0<%If CInt(GroupSet(2)) = 0 Then Response.Write " checked"%>> ��&nbsp;&nbsp;
      <input type=radio name="GroupSet(2)" value=1<%If CInt(GroupSet(2)) = 1 Then Response.Write " checked"%>> �� </td>
  </tr>
  <tr>
    <td class=tablerow2>�Ƿ����ʹ���ղؼ�</td>
    <td class=tablerow2><input type=radio name="GroupSet(3)" value=0<%If CInt(GroupSet(3)) = 0 Then Response.Write " checked"%>> ��&nbsp;&nbsp;
      <input type=radio name="GroupSet(3)" value=1<%If CInt(GroupSet(3)) = 1 Then Response.Write " checked"%>> �� </td>
  </tr>
  <tr>
    <td class=tablerow1>�Ƿ������Ӻ���</td>
    <td class=tablerow1><input type=radio name="GroupSet(4)" value=0<%If CInt(GroupSet(4)) = 0 Then Response.Write " checked"%>> ��&nbsp;&nbsp;
      <input type=radio name="GroupSet(4)" value=1<%If CInt(GroupSet(4)) = 1 Then Response.Write " checked"%>> �� </td>
  </tr>
  <tr>
    <td class=tablerow2>����ղض�������Ϣ -- ������������Ϊ0</td>
    <td class=tablerow2><input type=text name=GroupSet(5) size=10 value='<%=GroupSet(5)%>'></td>
  </tr>
  <tr>
    <td class=tablerow1>�����Ӷ��ٺ��� -- ������������Ϊ0</td>
    <td class=tablerow1><input type=text name=GroupSet(6) size=10 value='<%=GroupSet(6)%>'></td>
  </tr>
  <tr> 
    <th colspan="2" align=left>��������Ȩ������</th>
  </tr>
  <tr>
    <td class=tablerow1>���Է�������</td>
    <td class=tablerow1><input type=radio name="GroupSet(7)" value=0<%If CInt(GroupSet(7)) = 0 Then Response.Write " checked"%>> ��&nbsp;&nbsp;
      <input type=radio name="GroupSet(7)" value=1<%If CInt(GroupSet(7)) = 1 Then Response.Write " checked"%>> �� </td>
  </tr>

  <tr>
    <td class=tablerow2>���Է������ŵ�Ƶ��(35)��ʹ��$$$�ָĬ��1</td>
    <td class=tablerow2><input type=text name=GroupSet(35) size=10 value='<%=GroupSet(35)%>'></td>
  </tr>

  <tr>
    <td class=tablerow2>���Թ����Լ�����������</td>
    <td class=tablerow2><input type=radio name="GroupSet(8)" value=0<%If CInt(GroupSet(8)) = 0 Then Response.Write " checked"%>> ��&nbsp;&nbsp;
      <input type=radio name="GroupSet(8)" value=1<%If CInt(GroupSet(8)) = 1 Then Response.Write " checked"%>> �� </td>
  </tr>
  <tr>
    <td class=tablerow1>�����������ӵĵ���</td>
    <td class=tablerow1><input type=text name=GroupSet(9) size=10 value='<%=GroupSet(9)%>'></td>
  </tr>
  <tr>
    <td class=tablerow2>ÿ����Է�������ƪ����</td>
    <td class=tablerow2><input type=text name=GroupSet(10) size=10 value='<%=GroupSet(10)%>'></td>
  </tr>
  <tr>
    <td class=tablerow1>���Է������</td>
    <td class=tablerow1><input type=radio name="GroupSet(11)" value=0<%If CInt(GroupSet(11)) = 0 Then Response.Write " checked"%>> ��&nbsp;&nbsp;
      <input type=radio name="GroupSet(11)" value=1<%If CInt(GroupSet(11)) = 1 Then Response.Write " checked"%>> �� </td>
  </tr>
  <tr>
    <td class=tablerow2>���Թ����Լ����������</td>
    <td class=tablerow2><input type=radio name="GroupSet(12)" value=0<%If CInt(GroupSet(12)) = 0 Then Response.Write " checked"%>> ��&nbsp;&nbsp;
      <input type=radio name="GroupSet(12)" value=1<%If CInt(GroupSet(12)) = 1 Then Response.Write " checked"%>> �� </td>
  </tr>
  <tr>
    <td class=tablerow1>����������ӵĵ���</td>
    <td class=tablerow1><input type=text name=GroupSet(13) size=10 value='<%=GroupSet(13)%>'></td>
  </tr>
  <tr>
    <td class=tablerow2>ÿ����Է������ٸ����</td>
    <td class=tablerow2><input type=text name=GroupSet(14) size=10 value='<%=GroupSet(14)%>'></td>
  </tr>
  <tr>
    <td class=tablerow1>����������Ϣ��Ҫ����Ա���</td>
    <td class=tablerow1><input type=radio name="GroupSet(15)" value=0<%If CInt(GroupSet(15)) = 0 Then Response.Write " checked"%>> ��&nbsp;&nbsp;
      <input type=radio name="GroupSet(15)" value=1<%If CInt(GroupSet(15)) = 1 Then Response.Write " checked"%>> �� </td>
  </tr>
  <tr>
    <td class=tablerow2>�����������ݵ�����ֽ�</td>
    <td class=tablerow2><input type=text name=GroupSet(16) size=10 value='<%=GroupSet(16)%>'> byte</td>
  </tr>
  <tr>
    <td class=tablerow2>���������Ϣ������ֽ�</td>
    <td class=tablerow2><input type=text name=GroupSet(17) size=10 value='<%=GroupSet(17)%>'> byte</td>
  </tr>
  <tr>
    <td class=tablerow1>ɾ�����¿۳��ĵ���</td>
    <td class=tablerow1><input type=text name=GroupSet(18) size=10 value='<%=GroupSet(18)%>'></td>
  </tr>
  <tr>
    <td class=tablerow2>ɾ������۳��ĵ���</td>
    <td class=tablerow2><input type=text name=GroupSet(19) size=10 value='<%=GroupSet(19)%>'></td>
  </tr>
  <tr>
    <td class=tablerow1>�Ƿ�����ϴ��ļ�</td>
    <td class=tablerow1><input type=radio name="GroupSet(20)" value=0<%If CInt(GroupSet(20)) = 0 Then Response.Write " checked"%>> ��&nbsp;&nbsp;
      <input type=radio name="GroupSet(20)" value=1<%If CInt(GroupSet(20)) = 1 Then Response.Write " checked"%>> �� </td>
  </tr>
  <tr>
    <td class=tablerow2>ÿ������ϴ��ļ���</td>
    <td class=tablerow2><input type=text name=GroupSet(21) size=10 value='<%=GroupSet(21)%>'></td>
  </tr>
  <tr> 
    <th colspan="2" align=left>����վ�ڶ�������</th>
  </tr>
  <tr>
    <td class=tablerow1>�Ƿ���Է��Ͷ���</td>
    <td class=tablerow1><input type=radio name="GroupSet(22)" value=0<%If CInt(GroupSet(22)) = 0 Then Response.Write " checked"%>> ��&nbsp;&nbsp;
      <input type=radio name="GroupSet(22)" value=1<%If CInt(GroupSet(22)) = 1 Then Response.Write " checked"%>> �� </td>
  </tr>
  <tr>
    <td class=tablerow2>���Ͷ�����������</td>
    <td class=tablerow2><input type=text name=GroupSet(23) size=10 value='<%=GroupSet(23)%>'> byte</td>
  </tr>
  <tr>
    <td class=tablerow1>�����С���� -- ������������Ϊ0</td>
    <td class=tablerow1><input type=text name=GroupSet(24) size=10 value='<%=GroupSet(24)%>'> ��</td>
  </tr>
  <tr>
    <td class=tablerow2>ÿ����Է��Ͷ���������</td>
    <td class=tablerow2><input type=text name=GroupSet(29) size=10 value='<%=GroupSet(29)%>'></td>
  </tr>
  <tr> 
    <th colspan="2" align=left>������������</th>
  </tr>
  <tr>
    <td class=tablerow1>ÿ�ε�½���ӵĵ���</td>
    <td class=tablerow1><input type=text name=GroupSet(25) size=10 value='<%=GroupSet(25)%>'></td>
  </tr>
  <tr>
    <td class=tablerow2>ÿ�ε�½���Ӿ���ֵ</td>
    <td class=tablerow2><input type=text name=GroupSet(32) size=10 value='<%=GroupSet(32)%>'></td>
  </tr>
  <tr>
    <td class=tablerow1>ÿ�ε�½���ӵ�����ֵ</td>
    <td class=tablerow1><input type=text name=GroupSet(33) size=10 value='<%=GroupSet(33)%>'></td>
  </tr>
  <tr>
    <td class=tablerow2>������Ϣ���ӵĵ���</td>
    <td class=tablerow2><input type=text name=GroupSet(26) size=10 value='<%=GroupSet(26)%>'></td>
  </tr>
  <tr>
    <td class=tablerow1>�ظ��������ӵĵ���</td>
    <td class=tablerow1><input type=text name=GroupSet(27) size=10 value='<%=GroupSet(27)%>'></td>
  </tr>
  <tr>
    <td class=tablerow2>���������ܵ��ۿ�</td>
    <td class=tablerow2><input type=text name=GroupSet(28) size=5 value='<%=GroupSet(28)%>'> ��</td>
  </tr>
  <tr>
    <td class=tablerow1>�Ƿ���Թ���</td>
    <td class=tablerow1><input type=radio name="GroupSet(30)" value=0<%If CInt(GroupSet(30)) = 0 Then Response.Write " checked"%>> ��&nbsp;&nbsp;
      <input type=radio name="GroupSet(30)" value=1<%If CInt(GroupSet(30)) = 1 Then Response.Write " checked"%>> �� </td>
  </tr>
  <tr>
    <td class=tablerow2>�Ƿ�����������</td>
    <td class=tablerow2><input type=radio name="GroupSet(31)" value=0<%If CInt(GroupSet(31)) = 0 Then Response.Write " checked"%>> ��&nbsp;&nbsp;
      <input type=radio name="GroupSet(31)" value=1<%If CInt(GroupSet(31)) = 1 Then Response.Write " checked"%>> �� </td>
  </tr>
  <tr>
    <td class=tablerow1>��������Ƿ�ֱ����ʾ���ص�ַ(����Աʹ�ù�������)</td>
    <td class=tablerow1><input type=radio name="GroupSet(34)" value=0<%If CInt(GroupSet(34)) = 0 Then Response.Write " checked"%>> ��&nbsp;&nbsp;
      <input type=radio name="GroupSet(34)" value=1<%If CInt(GroupSet(34)) = 1 Then Response.Write " checked"%>> �� </td>
  </tr>
  <tr> 
    <td class=tablerow2></td>
    <td class=tablerow2>&nbsp;
      <input type="button" name="Submit1" onclick="javascript:history.go(-1)" value="������һҳ" class=button>��
      <input type="submit" name="submit" value="�����޸�" class=button></td>
  </tr>
  <input type=hidden value="<%=Request("groupid")%>" name="groupid">
  </form>
</table><br>	
<%
	   Rs.Close:Set Rs=Nothing
End Sub
Sub savegroup()
	If Len(request.form("GroupName")) = 0 Then
  		founderr=true
  		errmsg=errmsg+"<li>�û��鲻��Ϊ�գ�</li>"
		Exit Sub
	End If
	If Trim(request.form("Grades")) = "" Then
  		founderr = true
  		errmsg = errmsg+"<li>�û��ȼ�����Ϊ�գ�</li>"
		Exit Sub
	End If
	Set Rs = enchiasp.Execute("select Grades from [ECCMS_UserGroup] where Grades = " & CInt(Request("Grades")))
	If Not (Rs.bof And Rs.EOF) Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>Sorry������ͬ�ĵȼ�����,�����������û��ȼ����ԣ�</li>"
		Exit Sub
	End If
	Set Rs = enchiasp.Execute("select Groupset from [ECCMS_UserGroup] where Grades = 1")
	Groupsetting = enchiasp.CheckStr(Rs("Groupset"))
	Rs.Close:Set Rs = Nothing
	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL = "select * from [ECCMS_UserGroup] where (groupid is null)"
	Rs.Open SQL,Conn,1,3
	Rs.Addnew
		Rs("Groupid") = Request.Form("Newgroupid")
		Rs("Groupname") = Request.Form("Groupname")
		Rs("Grades") = Request.Form("Grades")
		Rs("Groupset") = Groupsetting
	Rs.Update
	Rs.Close:Set Rs=Nothing
	Succeed("<li>����û��� "&request.form("GroupName")&" �ɹ�!</li>")

End Sub
Sub savedit()
	If Len(request.form("GroupName")) = 0 Then
  		founderr=true
  		errmsg=errmsg+"<li>�û��鲻��Ϊ�գ�</li>"
		Exit Sub
	End If
	Dim Group_Setting
	For i = 0 To 35
		Group_Setting = Group_Setting & Request.Form("GroupSet(" & i & ")") & "|||"
	Next
	Group_Setting = Group_Setting & "0|||0|||0|||1|||1|||1|||0|||"
	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL = "select * from [ECCMS_UserGroup] where groupid = " & Request.Form("groupid")
	Rs.Open SQL,Conn,1,3
		Rs("Groupname") = Request.Form("Groupname")
		Rs("Grades") = Request.Form("Grades")
		Rs("Groupset") = enchiasp.Checkstr(Group_setting)
	Rs.Update
	Rs.Close:Set Rs=Nothing
	enchiasp.DelCahe "GroupSetting" & Request.Form("Grades")
	Succeed("<li>�޸��û��� "& Request.Form("GroupName") &" �ɹ�!</li>")
End Sub
Sub delgroup()
	enchiasp.DelCahe "GroupSetting" & Request.Form("Grade")
	enchiasp.execute("Delete From ECCMS_UserGroup where groupid="&request("groupid"))
	enchiasp.execute("update ECCMS_User set UserGrade=1 where UserGrade="&request("Grade"))
	Response.Redirect("admin_group.asp")
End Sub
%>
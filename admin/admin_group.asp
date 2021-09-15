<!--#include file="setup.asp"-->
<!--#include file="check.asp"-->
<%
Admin_header
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
<th height="23" colspan="4" >用户组管理&nbsp;&nbsp;|&nbsp;&nbsp;<a href="?action=addgroup"><font color=#FFFFFF><strong>[添加新用户组]</strong></font></a></th>
</tr>
<tr><td colspan=4 height=25 class="tablerow1"><B>说明</B>：<BR>
①在这里您可以设置各个用户组在系统中的默认权限，系统默认用户组不能删除和编辑用户等级;<BR>
②可以进行添加用户组操作并设置其权限，可以将其他组用户转移到该组，请到用户管理中进行相关操作;<BR>
③可以删除和编辑新添加的用户组，添加组时请填相应用户等级。<BR>
</td></tr>
<tr align=center>
<td height="23" width="30%" class=TableTitle><B>用户组</B></td>
<td height="23" width="20%" class=TableTitle><B>用户数量</B></td>
<td height="23" width="20%" class=TableTitle><B>编辑</B></td>
<td height="23" width="30%" class=TableTitle><B>用户等级</B></td>
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
<td height="23" <%=strClass%>><%if rs("Grades") = 0 Then%>匿名用户<%Else%><%=trs(0)%><%End If%></td>
<td height="23" <%=strClass%>><a href="?action=editgroup&groupid=<%=rs("groupid")%>">用户组设置</a><%if rs("groupid") => 6 then%> | <a href="?action=del&groupid=<%=rs("groupid")%>&Grade=<%=rs("Grades")%>" onclick="{if(confirm('此操作将删除本用户组\n 您确定执行的操作吗?')){return true;}return false;}">删除</a><%end if%></td>
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
<th height="23" colspan="2" >添加新的用户组</th>
</tr>
<FORM METHOD=POST ACTION="admin_group.asp?action=save">
<input type="hidden" name="newgroupid" value="<% = GroupNum %>">
<tr><td colspan=2 height=25 class="tablerow1"><B>说明</B>：<BR>
①可以进行添加用户组操作并设置其权限，可以将其他组用户转移到该组，请到用户管理中进行相关操作；<BR>
②可以删除和编辑新添加的用户组，添加是请填写相应用户等级。<BR>
</td></tr>
<tr> 
<th colspan="2" >添加新的用户组</th>
</tr>
<tr>
<td width="60%" class=tablerow1>用户组名称</td>
<td width="40%" class=tablerow1><input size=35 name="GroupName" type=text></td>
</tr>
<script>function checkclick(msg){if(confirm(msg)){event.returnValue=true;}else{event.returnValue=false;}}</script>
<tr>
<td class=tablerow1>用户组等级；请输入数字(数字越大级别越高)</td>
<td class=tablerow1><input size=10 name="Grades" type=text value=<%=conn.execute("Select max(Grades)from ECCMS_UserGroup where Grades <> 999")(0)+1%>></td>
</tr>
<tr> 
<td class=tablerow1>
</td>
<td class=tablerow1>
<input type="button" name="Submit1" onclick="javascript:history.go(-1)" value="返回上一页" class=button>　
<input type="submit" name="submit" value="添加用户组" class=button></td>
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
    <th colspan="2" >修改用户组</th>
  </tr>
  <form method=post action="admin_group.asp?action=savedit">
  <tr>
    <td colspan=2 height=25 class="tablerow1"><B>说明</B>：<BR>
    ①可以进行修改用户组操作并设置其权限，可以将其他组用户转移到该组，请到用户管理中进行相关操作；<BR>
    </td>
  </tr>
  <tr> 
    <th colspan="2">用户组设置</th>
  </tr>
  <tr>
    <td width="60%" class=tablerow1>用户组名称</td>
    <td width="40%" class=tablerow1><input size=35 name="GroupName" type=text value="<%=Rs("GroupName")%>"></td>
  </tr>
  <tr>
    <td class=tablerow2>用户组等级；请输入数字(<font color=blue>数字越大级别越高</font>)</td>
    <td class=tablerow2><input size=10 type=text value="<%=Rs("Grades")%>" disabled>
    <input size=10 name="Grades" type=hidden value="<%=Rs("Grades")%>">&nbsp;&nbsp;&nbsp;&nbsp;
    <a href="admin_group.asp">返回用户组首页</a></td>
  </tr>
  <tr> 
    <th colspan="2" align=left>＝＝用户基本使用设置</th>
  </tr>
  <tr>
    <td class=tablerow1>用户是否可以修改密码</td>
    <td class=tablerow1><input type=radio name="GroupSet(0)" value=0<%If CInt(GroupSet(0)) = 0 Then Response.Write " checked"%>> 否&nbsp;&nbsp;
      <input type=radio name="GroupSet(0)" value=1<%If CInt(GroupSet(0)) = 1 Then Response.Write " checked"%>> 是 </td>
  </tr>
  <tr>
    <td class=tablerow2>用户是否可以修改资料</td>
    <td class=tablerow2><input type=radio name="GroupSet(1)" value=0<%If CInt(GroupSet(1)) = 0 Then Response.Write " checked"%>> 否&nbsp;&nbsp;
      <input type=radio name="GroupSet(1)" value=1<%If CInt(GroupSet(1)) = 1 Then Response.Write " checked"%>> 是 </td>
  </tr>
  <tr>
    <td class=tablerow1>发布内容信息是否使用验证码</td>
    <td class=tablerow1><input type=radio name="GroupSet(2)" value=0<%If CInt(GroupSet(2)) = 0 Then Response.Write " checked"%>> 否&nbsp;&nbsp;
      <input type=radio name="GroupSet(2)" value=1<%If CInt(GroupSet(2)) = 1 Then Response.Write " checked"%>> 是 </td>
  </tr>
  <tr>
    <td class=tablerow2>是否可以使用收藏夹</td>
    <td class=tablerow2><input type=radio name="GroupSet(3)" value=0<%If CInt(GroupSet(3)) = 0 Then Response.Write " checked"%>> 否&nbsp;&nbsp;
      <input type=radio name="GroupSet(3)" value=1<%If CInt(GroupSet(3)) = 1 Then Response.Write " checked"%>> 是 </td>
  </tr>
  <tr>
    <td class=tablerow1>是否可以添加好友</td>
    <td class=tablerow1><input type=radio name="GroupSet(4)" value=0<%If CInt(GroupSet(4)) = 0 Then Response.Write " checked"%>> 否&nbsp;&nbsp;
      <input type=radio name="GroupSet(4)" value=1<%If CInt(GroupSet(4)) = 1 Then Response.Write " checked"%>> 是 </td>
  </tr>
  <tr>
    <td class=tablerow2>最多收藏多少条信息 -- 不限制请设置为0</td>
    <td class=tablerow2><input type=text name=GroupSet(5) size=10 value='<%=GroupSet(5)%>'></td>
  </tr>
  <tr>
    <td class=tablerow1>最多添加多少好友 -- 不限制请设置为0</td>
    <td class=tablerow1><input type=text name=GroupSet(6) size=10 value='<%=GroupSet(6)%>'></td>
  </tr>
  <tr> 
    <th colspan="2" align=left>＝＝发布权限设置</th>
  </tr>
  <tr>
    <td class=tablerow1>可以发布文章</td>
    <td class=tablerow1><input type=radio name="GroupSet(7)" value=0<%If CInt(GroupSet(7)) = 0 Then Response.Write " checked"%>> 否&nbsp;&nbsp;
      <input type=radio name="GroupSet(7)" value=1<%If CInt(GroupSet(7)) = 1 Then Response.Write " checked"%>> 是 </td>
  </tr>

  <tr>
    <td class=tablerow2>可以发布新闻的频道(35)，使用$$$分割，默认1</td>
    <td class=tablerow2><input type=text name=GroupSet(35) size=10 value='<%=GroupSet(35)%>'></td>
  </tr>

  <tr>
    <td class=tablerow2>可以管理自己发布的文章</td>
    <td class=tablerow2><input type=radio name="GroupSet(8)" value=0<%If CInt(GroupSet(8)) = 0 Then Response.Write " checked"%>> 否&nbsp;&nbsp;
      <input type=radio name="GroupSet(8)" value=1<%If CInt(GroupSet(8)) = 1 Then Response.Write " checked"%>> 是 </td>
  </tr>
  <tr>
    <td class=tablerow1>发布文章增加的点数</td>
    <td class=tablerow1><input type=text name=GroupSet(9) size=10 value='<%=GroupSet(9)%>'></td>
  </tr>
  <tr>
    <td class=tablerow2>每天可以发布多少篇文章</td>
    <td class=tablerow2><input type=text name=GroupSet(10) size=10 value='<%=GroupSet(10)%>'></td>
  </tr>
  <tr>
    <td class=tablerow1>可以发布软件</td>
    <td class=tablerow1><input type=radio name="GroupSet(11)" value=0<%If CInt(GroupSet(11)) = 0 Then Response.Write " checked"%>> 否&nbsp;&nbsp;
      <input type=radio name="GroupSet(11)" value=1<%If CInt(GroupSet(11)) = 1 Then Response.Write " checked"%>> 是 </td>
  </tr>
  <tr>
    <td class=tablerow2>可以管理自己发布的软件</td>
    <td class=tablerow2><input type=radio name="GroupSet(12)" value=0<%If CInt(GroupSet(12)) = 0 Then Response.Write " checked"%>> 否&nbsp;&nbsp;
      <input type=radio name="GroupSet(12)" value=1<%If CInt(GroupSet(12)) = 1 Then Response.Write " checked"%>> 是 </td>
  </tr>
  <tr>
    <td class=tablerow1>发布软件增加的点数</td>
    <td class=tablerow1><input type=text name=GroupSet(13) size=10 value='<%=GroupSet(13)%>'></td>
  </tr>
  <tr>
    <td class=tablerow2>每天可以发布多少个软件</td>
    <td class=tablerow2><input type=text name=GroupSet(14) size=10 value='<%=GroupSet(14)%>'></td>
  </tr>
  <tr>
    <td class=tablerow1>发布以上信息需要管理员审核</td>
    <td class=tablerow1><input type=radio name="GroupSet(15)" value=0<%If CInt(GroupSet(15)) = 0 Then Response.Write " checked"%>> 否&nbsp;&nbsp;
      <input type=radio name="GroupSet(15)" value=1<%If CInt(GroupSet(15)) = 1 Then Response.Write " checked"%>> 是 </td>
  </tr>
  <tr>
    <td class=tablerow2>发布文章内容的最多字节</td>
    <td class=tablerow2><input type=text name=GroupSet(16) size=10 value='<%=GroupSet(16)%>'> byte</td>
  </tr>
  <tr>
    <td class=tablerow2>发布软件信息的最多字节</td>
    <td class=tablerow2><input type=text name=GroupSet(17) size=10 value='<%=GroupSet(17)%>'> byte</td>
  </tr>
  <tr>
    <td class=tablerow1>删除文章扣除的点数</td>
    <td class=tablerow1><input type=text name=GroupSet(18) size=10 value='<%=GroupSet(18)%>'></td>
  </tr>
  <tr>
    <td class=tablerow2>删除软件扣除的点数</td>
    <td class=tablerow2><input type=text name=GroupSet(19) size=10 value='<%=GroupSet(19)%>'></td>
  </tr>
  <tr>
    <td class=tablerow1>是否可以上传文件</td>
    <td class=tablerow1><input type=radio name="GroupSet(20)" value=0<%If CInt(GroupSet(20)) = 0 Then Response.Write " checked"%>> 否&nbsp;&nbsp;
      <input type=radio name="GroupSet(20)" value=1<%If CInt(GroupSet(20)) = 1 Then Response.Write " checked"%>> 是 </td>
  </tr>
  <tr>
    <td class=tablerow2>每天可以上传文件数</td>
    <td class=tablerow2><input type=text name=GroupSet(21) size=10 value='<%=GroupSet(21)%>'></td>
  </tr>
  <tr> 
    <th colspan="2" align=left>＝＝站内短信设置</th>
  </tr>
  <tr>
    <td class=tablerow1>是否可以发送短信</td>
    <td class=tablerow1><input type=radio name="GroupSet(22)" value=0<%If CInt(GroupSet(22)) = 0 Then Response.Write " checked"%>> 否&nbsp;&nbsp;
      <input type=radio name="GroupSet(22)" value=1<%If CInt(GroupSet(22)) = 1 Then Response.Write " checked"%>> 是 </td>
  </tr>
  <tr>
    <td class=tablerow2>发送短信内容限制</td>
    <td class=tablerow2><input type=text name=GroupSet(23) size=10 value='<%=GroupSet(23)%>'> byte</td>
  </tr>
  <tr>
    <td class=tablerow1>信箱大小限制 -- 不限制请设置为0</td>
    <td class=tablerow1><input type=text name=GroupSet(24) size=10 value='<%=GroupSet(24)%>'> 条</td>
  </tr>
  <tr>
    <td class=tablerow2>每天可以发送多少条短信</td>
    <td class=tablerow2><input type=text name=GroupSet(29) size=10 value='<%=GroupSet(29)%>'></td>
  </tr>
  <tr> 
    <th colspan="2" align=left>＝＝其它设置</th>
  </tr>
  <tr>
    <td class=tablerow1>每次登陆增加的点数</td>
    <td class=tablerow1><input type=text name=GroupSet(25) size=10 value='<%=GroupSet(25)%>'></td>
  </tr>
  <tr>
    <td class=tablerow2>每次登陆增加经验值</td>
    <td class=tablerow2><input type=text name=GroupSet(32) size=10 value='<%=GroupSet(32)%>'></td>
  </tr>
  <tr>
    <td class=tablerow1>每次登陆增加的魅力值</td>
    <td class=tablerow1><input type=text name=GroupSet(33) size=10 value='<%=GroupSet(33)%>'></td>
  </tr>
  <tr>
    <td class=tablerow2>发布信息增加的点数</td>
    <td class=tablerow2><input type=text name=GroupSet(26) size=10 value='<%=GroupSet(26)%>'></td>
  </tr>
  <tr>
    <td class=tablerow1>回复留言增加的点数</td>
    <td class=tablerow1><input type=text name=GroupSet(27) size=10 value='<%=GroupSet(27)%>'></td>
  </tr>
  <tr>
    <td class=tablerow2>购物所享受的折扣</td>
    <td class=tablerow2><input type=text name=GroupSet(28) size=5 value='<%=GroupSet(28)%>'> 折</td>
  </tr>
  <tr>
    <td class=tablerow1>是否可以购物</td>
    <td class=tablerow1><input type=radio name="GroupSet(30)" value=0<%If CInt(GroupSet(30)) = 0 Then Response.Write " checked"%>> 否&nbsp;&nbsp;
      <input type=radio name="GroupSet(30)" value=1<%If CInt(GroupSet(30)) = 1 Then Response.Write " checked"%>> 是 </td>
  </tr>
  <tr>
    <td class=tablerow2>是否可以下载软件</td>
    <td class=tablerow2><input type=radio name="GroupSet(31)" value=0<%If CInt(GroupSet(31)) = 0 Then Response.Write " checked"%>> 否&nbsp;&nbsp;
      <input type=radio name="GroupSet(31)" value=1<%If CInt(GroupSet(31)) = 1 Then Response.Write " checked"%>> 是 </td>
  </tr>
  <tr>
    <td class=tablerow1>下载软件是否直接显示下载地址(供会员使用工具下载)</td>
    <td class=tablerow1><input type=radio name="GroupSet(34)" value=0<%If CInt(GroupSet(34)) = 0 Then Response.Write " checked"%>> 否&nbsp;&nbsp;
      <input type=radio name="GroupSet(34)" value=1<%If CInt(GroupSet(34)) = 1 Then Response.Write " checked"%>> 是 </td>
  </tr>
  <tr> 
    <td class=tablerow2></td>
    <td class=tablerow2>&nbsp;
      <input type="button" name="Submit1" onclick="javascript:history.go(-1)" value="返回上一页" class=button>　
      <input type="submit" name="submit" value="保存修改" class=button></td>
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
  		errmsg=errmsg+"<li>用户组不能为空！</li>"
		Exit Sub
	End If
	If Trim(request.form("Grades")) = "" Then
  		founderr = true
  		errmsg = errmsg+"<li>用户等级不能为空！</li>"
		Exit Sub
	End If
	Set Rs = enchiasp.Execute("select Grades from [ECCMS_UserGroup] where Grades = " & CInt(Request("Grades")))
	If Not (Rs.bof And Rs.EOF) Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>Sorry！有相同的等级存在,请重新输入用户等级再试！</li>"
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
	Succeed("<li>添加用户组 "&request.form("GroupName")&" 成功!</li>")

End Sub
Sub savedit()
	If Len(request.form("GroupName")) = 0 Then
  		founderr=true
  		errmsg=errmsg+"<li>用户组不能为空！</li>"
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
	Succeed("<li>修改用户组 "& Request.Form("GroupName") &" 成功!</li>")
End Sub
Sub delgroup()
	enchiasp.DelCahe "GroupSetting" & Request.Form("Grade")
	enchiasp.execute("Delete From ECCMS_UserGroup where groupid="&request("groupid"))
	enchiasp.execute("update ECCMS_User set UserGrade=1 where UserGrade="&request("Grade"))
	Response.Redirect("admin_group.asp")
End Sub
%>
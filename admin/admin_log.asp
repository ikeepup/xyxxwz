<!--#include file =setup.asp-->
<!--#include file =check.asp-->
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
	<TH colspan=5>后台日志管理</TH>
</TR>
<TR height=25>
	<TD colspan=5 class=TableRow2> <B>选择查看日志事件：</B> 
	<A HREF=admin_log.asp>查看全部日志</A> | <A HREF=admin_log.asp?type=0>查看事件0</A> | <A HREF=admin_log.asp?type=1>查看事件1</A></TD>        
</TR>
<TR>
	<TH noWrap>操作</TH>
	<TH noWrap>操 作 人</TH>
	<TH noWrap> 对 象 </TH>
	<TH width='70%'>事件内容</TH>
	<TH noWrap>日期时间/IP</TH>
</TR><form action=admin_log.asp?action=del&LogType= method=post name=even>
<%
	Set Rs = Server.CreateObject("ADODB.Recordset")
	maxperpage = 20 '###每页显示数
	If Not IsNumeric(Request("page")) And Len(Request("page")) <> 0 Then
		Response.Write ("错误的系统参数!请输入整数")
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
	TotalPageNum = CInt(TotalNumber / maxperpage)  '得到总页数
	If TotalPageNum < TotalNumber / maxperpage Then TotalPageNum = TotalPageNum + 1
	If CurrentPage < 1 Then CurrentPage = 1
	If CurrentPage > TotalPageNum Then CurrentPage = TotalPageNum
	SQL = "select * from ECCMS_LogInfo "& WhereSQL &" order by logid desc"
	Rs.Open SQL, lConn, 1, 1
	If Rs.bof And Rs.EOF Then
		Response.Write "<tr><td colspan=5 class=TableRow1>还没有找到任何日志！</td></tr>"
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
	<TD class=TableRow1 align=center colspan=5><B>请选择要删除的日志事件:<B> 
	<input type=checkbox name=chkall value=on onclick="CheckAll(this.form)"> 全选         
	 <input class=Button type=submit name=act value=删除日志  onclick="{if(confirm('您确定执行的操作吗?')){this.document.even.submit();return true;}return false;}">
	 <input class=Button type=submit name=act onclick="{if(confirm('确定清除所有的日志纪录吗?')){this.document.even.submit();return true;}return false;}" value=清空日志>        
	 <input class=Button type=submit name=act onclick="{if(confirm('确定压缩日志数据库吗?')){this.document.even.submit();return true;}return false;}" value=压缩数据库></TD>        
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
			Response.Write "共有日志 <font COLOR=#FF0000><strong>" & totalnumber & "</strong></font> 条&nbsp;首 页&nbsp;上一页&nbsp;"
		Else
			Response.Write "共有日志 <font COLOR=#FF0000><strong>" & totalnumber & "</strong></font> 条&nbsp;<a href=?page=1&type=" & Request("calsid") & ">首 页</a>&nbsp;"
			Response.Write "<a href=?page=" & CurrentPage - 1 & "&type=" & Request("type") & ">上一页</a>&nbsp;"
		End If
		If n - CurrentPage < 1 Then
			Response.Write "下一页&nbsp;尾 页" & vbCrLf
		Else
			Response.Write "<a href=?page=" & (CurrentPage + 1) & "&type=" & Request("type") & ">下一页</a>"
			Response.Write "&nbsp;<a href=?page=" & n & "&type=" & Request("type") & ">尾 页</a>" & vbCrLf
		End If
		Response.Write "&nbsp;页次：<strong><font color=red>" & CurrentPage & "</font>/" & n & "</strong>页 "
		Response.Write "&nbsp;转到："
			Response.Write "&nbsp;<select name='page' size='1' style=""font-size: 9pt"" onChange='javascript:submit()'>" & vbCrLf
			For ii = 1 To n
				Response.Write "<option value='" & ii & "' "
				If CurrentPage = CInt(ii) Then
					Response.Write "selected "
				End If
				Response.Write ">第" & ii & "页</option>"
			Next
			Response.Write "</select> " & vbCrLf
		Response.Write "</td></tr></FORM></table>" & vbCrLf
End Sub

Sub BatchDel()
	Dim logid
	If Request("act")="删除日志" Then
		If request.form("logid")="" Then
			ErrMsg =  "请指定相关事件。"
			Founderr = True
			Exit Sub
		Else
			logid=replace(Request.Form("logid"),"'","")
			logid=replace(logid,";","")
			logid=replace(logid,"--","")
			logid=replace(logid,")","")
		End If
	End If
	If Request("act")="压缩数据库" Then
		If CompressMDB("Logdata.Asa") Then OutHintScript("恭喜您 ^_^ 日志数据库压缩成功！")
	ElseIf Request("act")="删除日志" Then
			lConn.Execute("delete from ECCMS_LogInfo where Datediff('D',LogAddTime, Now()) > 3 And logid in ("&logid&")")
	ElseIf Request("act")="清空日志" Then
		If request("LogType")="" or IsNull(Request("LogType")) Then 
			lConn.Execute("delete from ECCMS_LogInfo Where Datediff('D',LogAddTime, Now) > 3")
		Else
			lConn.Execute("delete from ECCMS_LogInfo where  Datediff('D',LogAddTime, Now()) > 3 And LogType="&CInt(request("LogType"))&"")
		End If
	End If
	Succeed ("成功删除日志。注意：三天内的日志会被系统保留。")
End Sub
'================================================
' 函数名：CompressMDB
' 作  用：压缩ACCESS数据库
' 参  数：dbPath ----数据库路径
' 返回值：True  ----  False
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
<!--#include file="config.asp"-->
<!--#include file="check.asp"-->
<!--#include file="../inc/ubbcode.asp"-->
<!--#include file="head.inc"-->
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
Call InnerLocation("发布软件")
If CInt(GroupSetting(11)) = 0 Then
	ErrMsg = ErrMsg + "<li>对不起！您的会员权限不足，如需要该权限请联系管理员。</li>"
	Founderr = True
End If


if CInt(enchiasp.membergrade)<CInt(enchiasp.postgrade) then
	ErrMsg = ErrMsg + "<li>对不起！您的权限不足,您没有发布软件的权限，如需要该权限请联系管理员。</li>"
	Founderr = True

end if


Dim Rs,SQL,i,SoftID
ChannelID = enchiasp.ChkNumeric(Request("ChannelID"))

if Request("ChannelID")="" then
	ErrMsg = ErrMsg + "<li>参数错误,请不要手工设置参数。</li>"
	Founderr = True
end if

If ChannelID < 2 Then ChannelID = 2
ChannelID = CLng(ChannelID)

Dim Action:Action = LCase(Request("action"))
Select Case Trim(Action)
	Case "save"
		Call SaveNewSoft
	Case "view"
		Call SoftView
	Case Else
		Call showmain
End Select
If Founderr = True Then
	Call Returnerr(ErrMsg)
End If

Sub showmain()
	If Founderr = True Then Exit Sub
	Dim Channel_Setting
	Channel_Setting = Split(enchiasp.Channel_Setting, "|||")
%>
<script language=JavaScript>
var _maxCount = '<%=CLng(GroupSetting(17))%>';
function ToRunsystem(addTitle) {
	var revisedTitle;
	var currentTitle;
	currentTitle = document.myform.RunSystem.value;
	revisedTitle = currentTitle+addTitle;
	document.myform.RunSystem.value=revisedTitle;
	document.myform.RunSystem.focus();
	return; 
}
function doSubmit(){
	if (document.myform.SoftName.value==""){
		alert("软件名称不能为空！");
		return false;
	}
	if (document.myform.DownAddress.value==""){
		alert("最起码要填写一个下载地址吧！");
		return false;
	}
	if (document.myform.SiteName.value==""){
		alert("下载名称不能为空！");
		return false;
	}
	if (document.myform.ClassID.value==""){
		alert("该一级分类已经有下属分类，请选择其下属分类！");
		return false;
	}
	if (document.myform.ClassID.value=="0"){
		alert("该分类是外部连接，不能添加内容！");
		return false;
	}
	if (document.myform.RunSystem.value==""){
		alert("软件运行环境不能为空！");
		return false;
	}
	if (document.myform.SoftType.value==""){
		alert("软件类型不能为空！");
		return false;
	}
	if (document.myform.SoftSize.value==""){
		alert("软件大小还没有填写！");
		return false;
	}
	<%If CInt(GroupSetting(2)) = 1 Then%>
	if (document.myform.codestr.value==""){
		alert("请填写验证码！");
		return false;
	}
	<%End If%>
	myform.content1.value = getHTML(); 
	MessageLength = Composition.document.body.innerHTML.length;
	if(MessageLength < 2){
		alert("软件简介不能小于2个字符！");
		return false;
	}
	if(MessageLength > _maxCount){
		alert("软件简介不能超过"+_maxCount+"个字符！");
		return false;
	}
	document.myform.Submit1.disabled = true;
	document.myform.submit();
}
</script>
<div onkeydown=CtrlEnter()>
<table  border="0" align="center" cellpadding="3" cellspacing="1" class="UserTableBorder">
        <tr>
          <th colspan="4">&gt;&gt;发布软件&lt;&lt;</th>
        </tr>
	<form method=Post name="myform" action="softpost.asp">
	<input type="Hidden" name="action" value="save">
	<input type=hidden name=ChannelID value="<%=ChannelID%>">
        <tr>
          <td width="15%" align="right" nowrap class="UserTableRow2"><strong>所属分类</strong></td>
          <td width="85%" class="UserTableRow1">
<%
	Dim sClassSelect
	Response.Write "<select name=""ClassID"" id=""ClassID"">"
	sClassSelect = enchiasp.LoadSelectClass(ChannelID)
	Response.Write sClassSelect
	Response.Write "</select>"
%>
	  </td>
        </tr>
        <tr>
          <td align="right" class="UserTableRow2"><strong>软件名称</strong></td>
          <td class="UserTableRow1"><input name="SoftName" type="text" id="SoftName" size="45" value=""> 
          <font color=red>*</font> <strong>软件版本</strong><input name="SoftVer" type="text" id="SoftVer" size="20" value=""></td>
	  </tr>
        <tr>
          <td align="right" class="UserTableRow2"><strong>相关软件</strong></td>
          <td class="UserTableRow1"><input name="Related" type="text" id="Related" size="60" value=""> <font color=red>*</font></td>
        </tr>
        <tr>
          <td align="right" class="UserTableRow2"><strong>运行环境</strong></td>
          <td class="UserTableRow1"><input name="RunSystem" type="text" size="60" value="<%=Channel_Setting(1)%>"><br>
<%
	Dim RunSystem
	RunSystem = Split(Channel_Setting(0), "|")
	For i = 0 To UBound(RunSystem)
		Response.Write "<a href='javascript:ToRunsystem(""" & Trim(RunSystem(i)) & """)'><u>" & Trim(RunSystem(i)) & "</u></a>  "
		If i = 10 Then Response.Write "<br>"
	Next
%>
          </td>
        </tr>
	<tr>
          <td align="right" class="UserTableRow2"><strong>软件类型</strong></td>
          <td class="UserTableRow1">
<%
	Dim SoftType
	SoftType = Split(Channel_Setting(2), ",")
	For i = 0 To UBound(SoftType)
		Response.Write "<input type=""radio"" name=""SoftType"" value=""" & Trim(SoftType(i)) & """ "
		If i = 0 Then Response.Write " checked"
		Response.Write ">" & Trim(SoftType(i)) & " "
		If i = 6 Then Response.Write "<br>"
	Next
%>
	  </td>
        </tr>
        <tr>
          <td align="right" class="UserTableRow2"><strong>预览图片</strong></td>
          <td class="UserTableRow1"><input name="SoftImage" id="ImageUrl" type="text" size="60" value=""></td>
        </tr>
	<tr>
          <td align="right" class="UserTableRow2"><strong>上传图片</strong></td>
          <td class="UserTableRow1"><iframe name="image" frameborder=0 width='100%' height=55 scrolling=no src=upload.asp?ChannelID=2></iframe></td>
        </tr>
	<tr>
          <td align="right" class="UserTableRow2"><strong>软件大小</strong></td>
          <td class="UserTableRow1">
	<input type="text" name="SoftSize" size="14" id="filesize" onkeyup="if(isNaN(this.value))this.value=''" value=''> <input name="SizeUnit" type="radio" value="KB" checked> KB <input type="radio" name="SizeUnit" value="MB"> MB <font color="#FF0000">！</font>
	<strong>解压密码</strong>
	<input type="text" name="Decode" size="15" maxlength="100" value=''> <font color="#808080">没有请留空</font>
          </td>
	</tr>
	<tr>
          <td align="right" class="UserTableRow2"><strong>软件性质</strong></td>
          <td class="UserTableRow1">
<%
	Response.Write " <select name=""impower"">"
	Dim ImpowerStr
	ImpowerStr = Split(Channel_Setting(3), ",")
	For i = 0 To UBound(ImpowerStr)
		Response.Write " <option value=""" & ImpowerStr(i) & """>" & ImpowerStr(i) & "</option>"
	Next
	Response.Write " </select>&nbsp;&nbsp;"
	Response.Write " <select name=""Languages"">"
	Response.Write " "
	Dim LanguagesStr
	LanguagesStr = Split(Channel_Setting(4), ",")
	For i = 0 To UBound(LanguagesStr)
		Response.Write " <option value=""" & LanguagesStr(i) & """>" & LanguagesStr(i) & "</option>"
	Next
	Response.Write " </select>&nbsp;&nbsp;"
%>
		<select name="star">
		<option value=5>★★★★★</option>
          	<option value=4>★★★★</option>
          	<option value=3 selected>★★★</option>
		<option value=2>★★</option>
		<option value=1>★</option>
          </select>&nbsp;&nbsp;
	  <strong><font color=blue>注册软件的价格</font></strong>
	  <input name="SoftPrice" type="text" size="10" onkeyup="if(isNaN(this.value))this.value=''" value="0"> 元
	  </td>
	</tr>
	<tr>
          <td align="right" class="UserTableRow2"><strong>联系方式</strong></td>
          <td class="UserTableRow1">
		<input name="Contact" type="text" size="33"> 
		<strong>程序主页</strong>
		<input name="Homepage" type="text" size="30">
	  </td>
	</tr>
	<tr>
          <td align="right" class="UserTableRow2"><strong>软件作者</strong></td>
          <td class="UserTableRow1">
		<input name="Author" type="text" size="33">
		<strong>注册网址</strong>
		<input name="Regsite" type="text" size="30">
	  </td>
        <tr>
          <td align="right" class="UserTableRow2"><strong>软件简介</strong></td>
          <td class="UserTableRow1"><textarea name='content1' id='content1' style='display:none'></textarea>
		<script Language=Javascript src="../editor/editor1.js"></script></td>
        </tr>
	        </tr>
<%
	If CInt(GroupSetting(2)) = 1 Then
%>
	<tr>
		<td class=UserTableRow2 align="right"><strong>验证码</strong></td>
		<td class=UserTableRow1><input type="text" name="codestr" maxlength="4" size="4">&nbsp;<img src="../inc/getcode.asp"></td>
	</tr>
<%
	End If
%>
	<tr>
          <td align="right" class="UserTableRow2"><strong>禁止评论</strong></td>
          <td class="UserTableRow1"><input name="ForbidEssay" type="checkbox" id="ForbidEssay" value="1">&nbsp;&nbsp;&nbsp;&nbsp;
	  <strong>立即发布</strong>
	  <input name="isAccept" type="checkbox" id="isAccept" value="1" disabled> 是（<font color=blue>如果选中的话将直接发布，否则审核后才能发布。</font>）</td>
        </tr>
	<tr>
          <td align="right" class="UserTableRow2"><strong>下载地址</strong></td>
          <td class="UserTableRow1"><input name="SiteName" type="text" size="15" value="下载地址1">
	  <input name="DownAddress" type="text" id="filePath" size="60" value=""> <font color=red>*</font></td>
        </tr>
	<tr>
          <td align="right" class="UserTableRow2"><strong>下载地址</strong></td>
          <td class="UserTableRow1"><input name="SiteName" type="text" size="15" value="下载地址2">
	  <input name="DownAddress" type="text" size="60" value=""></td>
        </tr>
	<tr>
          <td align="right" class="UserTableRow2"><strong>下载地址</strong></td>
          <td class="UserTableRow1"><input name="SiteName" type="text" size="15" value="下载地址3">
	  <input name="DownAddress" type="text" size="60" value=""></td>
        </tr>
	<tr>
          <td align="right" class="UserTableRow2"><strong>文件上传</strong></td>
          <td class="UserTableRow1"><iframe name="file1" frameborder=0 width='100%' height=60 scrolling=no src=upfile.asp?ChannelID=<%=ChannelID%>></iframe></td>
        </tr>
        <tr align="center">
          <td colspan="4" class="UserTableRow2">
	  <input type="button" name="Submit4" onclick="javascript:history.go(-1)" value="返回上一页" class=Button>
	  <input type="button" name="Submit1" value="现在发布" class=Button onclick="doSubmit();"></td>
        </tr></form>
      </table></div>
<%
End Sub

Function FormatDownAddress(ByVal str)
	If Trim(str) = ""  Or Trim(str) = "|||" Then
		FormatDownAddress = ""
		Exit Function
	End If
	Dim strDownAddress,sDownAddress,sDownSiteName
	Dim n,AddressNum,strAddress,strDownName,strTemp
	On Error Resume Next
	strDownAddress = Split(str, "|||")
	sDownAddress = Split(strDownAddress(1), "|")
	sDownSiteName = Split(strDownAddress(0), "|")
	If UBound(sDownAddress) < UBound(sDownSiteName) Then
		AddressNum = UBound(sDownAddress)
	Else
		AddressNum = UBound(sDownSiteName)
	End If
	strAddress = ""
	strDownName = ""
	For n = 0 To CInt(AddressNum)
		If Trim(sDownAddress(n)) <> "" And Trim(sDownSiteName(n)) <> "" Then
			strAddress = strAddress & Trim(sDownAddress(n)) & "|"
			strDownName = strDownName & Trim(sDownSiteName(n)) & "|"
		End If
	Next
	If Len(strDownName) > 0 Then strDownName = Left(strDownName, Len(strDownName) - 1)
	If Len(strAddress) > 0 Then strAddress = Left(strAddress, Len(strAddress) - 1)
	strTemp = strDownName & "|||" & strAddress
	FormatDownAddress = Trim(strTemp)
End Function
Sub SaveNewSoft()
	Dim TextContent,isAccept,ForbidEssay,DownAddress
	If CLng(UserToday(2)) => CLng(GroupSetting(14)) Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>您每天最多只能发布<font color=red><b>" & GroupSetting(14) & "</b></font>个软件，如果还要继续发布请明天再来吧！</li>"
	End If
	'禁止外部提交
	If enchiasp.CheckPost=false Then
		ErrMsg = ErrMsg + Postmsg
		FoundErr = True
		Exit Sub
	End If
	If CInt(GroupSetting(2)) = 1 Then
		If Not enchiasp.CodeIsTrue() Then
			ErrMsg = ErrMsg + "<meta http-equiv=refresh content=""2;URL="&Request.ServerVariables("HTTP_REFERER")&"""><li>验证码校验失败，请返回刷新页面再试。两秒后自动返回</li>"
			Founderr = True
			Exit Sub
		End If
		Session("GetCode") = ""
	End If

	If Trim(Request.Form("SoftName")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>软件名称不能为空！</li>"
	End If
	If Len(Request.Form("SoftName")) => 200 Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>软件名称不能超过200个字符！</li>"
	End If
	If Len(Request.Form("Related")) => 200 Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>相关软件不能超过200个字符！</li>"
	End If
	If Not IsNumeric(Request.Form("star")) Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>软件星级不能为空。</li>"
	End If
	If CLng(Request.Form("ClassID")) = 0 Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>该分类是外部连接，不能添加软件！</li>"
	End If
	If Not IsNumeric(Request.Form("ClassID")) Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>该一级分类已经有下属分类，不能添加软件！</li>"
	End If
	If Trim(Request.Form("SoftType")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>请选择软件类型！</li>"
	End If
	If Trim(Request.Form("impower")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>请选择软件授权方式！</li>"
	End If
	If Trim(Request.Form("Languages")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>请选择软件语言！</li>"
	End If
	If Trim(Request.Form("content1")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>软件简介不能为空！</li>"
	End If
	TextContent = ""
	For i = 1 To Request.Form("content1").Count
		TextContent = TextContent & Request.Form("content1")(i)
	Next
	If Len(Request.Form("RunSystem")) = 0 Then
		Founderr = True
		ErrMsg = ErrMsg + "<li>运行环境不能为空！</li>"
	End If
	If Not IsNumeric(Request.Form("SoftSize")) Then
		Founderr = True
		ErrMsg = ErrMsg + "<li>软件大小请输入整数！</li>"
	End If
	If CInt(Request("isAccept")) = 1 Then
		isAccept = 1
	Else
		isAccept = 0
	End If
	If CInt(Request.Form("ForbidEssay")) = 1 Then
		ForbidEssay = 1
	Else
		ForbidEssay = 0
	End If
	'---- 提取下载地址表单中的数据
	Dim TempAddress,TempSiteName,TempDownAddress
	Dim strTempAddress,strTempSiteName
	If Trim(Request.Form("DownAddress")) <> "" And Trim(Request.Form("SiteName")) <> "" Then
		strTempAddress = ""
		For Each TempAddress In Request.Form("DownAddress")
			If LCase(TempAddress) <> "del" And Trim(TempAddress) <> "" Then
				strTempAddress = strTempAddress & Replace(TempAddress, "|", "") & "|"
			End If
		Next
		If Len(strTempAddress) > 0 Then strTempAddress = Left(strTempAddress, Len(strTempAddress) - 1)
		strTempSiteName = ""
		For Each TempSiteName In Request.Form("SiteName")
			If LCase(TempSiteName) <> "del" And Trim(TempSiteName) <> "" Then
				strTempSiteName = strTempSiteName & Replace(TempSiteName, "|", "") & "|"
			End If
		Next
		If Len(strTempSiteName) > 0 Then strTempSiteName = Left(strTempSiteName, Len(strTempSiteName) - 1)
		TempDownAddress = enchiasp.CheckStr(strTempSiteName &"|||"& strTempAddress)
	Else
		TempDownAddress = ""
	End If
	DownAddress = FormatDownAddress(TempDownAddress)
	If Founderr = True Then Exit Sub
	On Error Resume Next
	Call PreventRefresh  '防刷新
	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL = "SELECT * FROM ECCMS_SoftList WHERE (SoftID is null)"
	Rs.Open SQL,Conn,1,3
	Rs.Addnew
		Rs("ChannelID") = ChannelID
		Rs("ClassID") = enchiasp.ChkNumeric(Request.Form("ClassID"))
		Rs("SpecialID") = 0
		Rs("SoftName") = enchiasp.ChkFormStr(Request.Form("SoftName"))
		Rs("SoftVer") = enchiasp.ChkFormStr(Request.Form("SoftVer"))
		Rs("ColorMode") = 0
		Rs("FontMode") = 0
		Rs("Related") = enchiasp.ChkFormStr(Request.Form("Related"))
		Rs("Content") = Html2Ubb(TextContent)
		Rs("Languages") = enchiasp.ChkFormStr(Request.Form("Languages"))
		Rs("SoftType") = enchiasp.ChkFormStr(Request.Form("SoftType"))
		Rs("RunSystem") = enchiasp.ChkFormStr(Request.Form("RunSystem"))
		Rs("impower") = enchiasp.ChkFormStr(Request.Form("impower"))
		If UCase(Request.Form("SizeUnit")) = "MB" Then
			Rs("SoftSize") = enchiasp.CheckNumeric(Request.Form("SoftSize") * 1024)
		Else
			Rs("SoftSize") = enchiasp.CheckNumeric(Request.Form("SoftSize"))
		End If
		Rs("star") = enchiasp.ChkNumeric(Request.Form("star"))
		Rs("Homepage") = enchiasp.ChkFormStr(Request.Form("Homepage"))
		Rs("Contact") = enchiasp.ChkFormStr(Request.Form("Contact"))
		Rs("Author") = enchiasp.ChkFormStr(Request.Form("Author"))
		Rs("Regsite") = enchiasp.ChkFormStr(Request.Form("Regsite"))
		Rs("showreg") = 0
		Rs("username") = Trim(enchiasp.MemberName)
		Rs("PointNum") = 0
		Rs("SoftPrice") = enchiasp.CheckNumeric(Request.Form("SoftPrice"))
		Rs("SoftTime") = Now()
		Rs("isTop") = 0
		Rs("AllHits") = 0
		Rs("DayHits") = 0
		Rs("WeekHits") = 0
		Rs("MonthHits") = 0
		Rs("HitsTime") = Now()
		Rs("HtmlFileDate") = Trim(enchiasp.HtmlRndFileName)
		Rs("SoftImage") = enchiasp.ChkFormStr(Request.Form("SoftImage"))
		Rs("Decode") = enchiasp.ChkFormStr(Request.Form("Decode"))
		Rs("isBest") = 0
		Rs("UserGroup") = 0
		Rs("isUpdate") = 1
		Rs("ErrCode") = 0
		Rs("downid") = 0
		Rs("DownFileName") = ""
		Rs("DownAddress") = enchiasp.ChkFormStr(DownAddress)
		Rs("isAccept") = isAccept
		Rs("ForbidEssay") = ForbidEssay
	Rs.update
	Rs.Close
	Rs.Open "select top 1 softid from ECCMS_SoftList where ChannelID=" & ChannelID & " order by softid desc", Conn, 1, 1
	SoftID = Rs("SoftID")
	Rs.Close:Set Rs = Nothing
	Dim strUserToday
	strUserToday = UserToday(0) &","& UserToday(1) &","& UserToday(2)+1 &","& UserToday(3) &","& UserToday(4) &","& UserToday(5)
	UpdateUserToday(strUserToday)
	Call Returnsuc("<li>恭喜您！提交成功。请等待管理员验证后正式发布。</li><li><a href=?action=view&ChannelID=" & ChannelID & "&SoftID=" & SoftID & ">点击此处查看该软件</a></li>")
End Sub
Private Sub SoftView()
	If Request("SoftID") = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>Sorry！您选择了错误的系统参数。</li>"
		Exit Sub
	End If
	If ChannelID = 0 Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>请指定频道。</li>"
		Exit Sub
	End If
	SQL = "SELECT * FROM ECCMS_SoftList WHERE ChannelID=" & ChannelID & " And username='" & enchiasp.MemberName & "' And SoftID=" & Request("SoftID")
	Set Rs = enchiasp.Execute(SQL)
	If Rs.bof And Rs.EOF Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>Sorry！没有找到任何软件。或者您选择了错误的系统参数！</li>"
		Exit Sub
	Else
	Dim strDownAddress,sDownAddress
	strDownAddress = Split(Rs("DownAddress"), "|||")
	sDownAddress = Split(strDownAddress(1), "|")
%>
<table border="0" align="center" cellpadding="3" cellspacing="1" class="UserTableBorder" style="table-layout:fixed;word-break:break-all">
	<tr>
	  <th colspan="2">&gt;&gt;查看软件信息&lt;&lt;</th>
	</tr>
	<tr>
	  <td align="center" class="UserTableRow2" colspan="2"><font size=3 color=blue><%=enchiasp.ReadFontMode(Rs("SoftName"),Rs("ColorMode"),Rs("FontMode"))%>&nbsp;<%=Rs("SoftVer")%></font></td>
	</tr>
	<tr>
	  <td class="UserTableRow1"><strong>软件运行环境：</strong> <%=Rs("RunSystem")%></td>
	  <td class="UserTableRow1"><strong>软件类型：</strong> <%=Rs("SoftType")%></td>
	</tr>
	<tr>
	  <td class="UserTableRow1"><strong>软件大小：</strong> <%=Rs("SoftSize")%></td>
	  <td class="UserTableRow1"><strong>软件星级：</strong> 
<%
Response.Write "<font color=red>"
For i = 1 to Rs("star")
	Response.Write "★"
Next
Response.Write "</font>"
%>
	  </td>
	</tr>
	<tr>
	  <td class="UserTableRow1"><strong>软件语言：</strong> <%=Rs("Languages")%></td>
	  <td class="UserTableRow1"><strong>授权方式：</strong> <%=Rs("impower")%></td>
	</tr>
	<tr>
	  <td class="UserTableRow1"><strong>更新时间：</strong> <%=Rs("SoftTime")%></td>
	  <td class="UserTableRow1"><strong>程序主页：</strong> <%=Rs("Homepage")%></td>
	</tr>
	<tr>
	  <td colspan="2" class="UserTableRow1"><strong>软件简介：</strong><br><%=UBBCode(Rs("content"))%></td>
	</tr>
	<tr>
	  <td colspan="2" class="UserTableRow1"><strong>下载地址：</strong><br>
<%
	For i = 0 To UBound(sDownAddress)
		Response.Write "<li><a href=""" & sDownAddress(i) & """ target=_blank>" & sDownAddress(i) & "</a></li>" & vbNewLine
	Next

%>
	  </td>
	</tr>
	<tr>
	  <td align="center" colspan="2" class="UserTableRow2"><input type="button" onclick="javascript:window.close()" value="关闭本窗口" name="B2" class=Button>&nbsp;&nbsp;
	  <input type="button" name="Submit1" onclick="javascript:location.href='#'" value="返回顶部" class=button>&nbsp;&nbsp;
	  </td>
	</tr>
</table>
<%

	End If
	Rs.Close
	Set Rs = Nothing 
End Sub

%><!--#include file="foot.inc"-->
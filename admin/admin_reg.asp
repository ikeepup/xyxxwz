<!--#include file="setup.asp" -->
<!--#include file="check.asp"-->

<%
'=====================================================================
' 软件名称：恩池网站管理系统--注册管理
' 当前版本：enchicms Version 3.0.0
' 更新日期：2005-03-25
' 官方网站：运城市恩池软件科技开发有限公司(www.enchi.com.cn) 
' 技术支持：柳云帆
' 邮箱：liuyunfan@163.com
' QQ：21556923
'=====================================================================
' Copyright 2005-2008  All Rights Reserved.
'=====================================================================
Dim selAdminID
Dim i,Action,strClass
Admin_header
If Not ChkAdmin("999") Then
	Server.Transfer("showerr.asp")
	Response.End
End If
Response.Write "<table cellpadding=2 cellspacing=1 border=0 class=tableBorder align=center>" & vbCrLf
Response.Write " <tr>" & vbCrLf
Response.Write " <th height=22 colspan=6>软件注册</th>" & vbCrLf
Response.Write " </tr>" & vbCrLf
Response.Write " <tr>" & vbCrLf
Response.Write " <td class=TableRow1> <b>备注：</b> 为保护知识产权，为更好获得系统服务，建议您进行软件的产品注册，否则将无法获取相关的升级服务，对由此而产生的任何故障将得不到技术上的支持。以下各项内容为必填项，请认真填写，如有不清楚地方请咨询公司技术人员。谢谢合作！"
Response.Write " </td>" & vbCrLf
Response.Write " </tr>" & vbCrLf
Response.Write " </table><br>" & vbCrLf


Action = LCase(Request("action"))
Select Case Trim(Action)
Case "reg"
	Call savereg
Case Else
	Call reginfo
End Select
If FoundErr = True Then
	ReturnError(ErrMsg)
End If
Admin_footer
SaveLogInfo(AdminName)
CloseConn

Private Sub reginfo()
	dim urlflag
	Response.Write "<table cellpadding=""2"" cellspacing=""1"" border=""0"" class=""tableBorder"" align=center>" & vbCrLf
	Response.Write "<form action='?action=reg' method=post>" & vbCrLf

	Response.Write " <tr>" & vbCrLf
	Response.Write " <th height=22 colspan=2>注册选项</th>" & vbCrLf
	Response.Write " </tr>" & vbCrLf
	Set Rs = enchiasp.Execute("select * from ECCMS_config")
	if rs.eof then
		response.write "数据库配置出错，请检查！"
	else
		
		urlflag=rs("urlflag")
		Response.Write " <tr>" & vbCrLf
		Response.Write "<td Class=TableRow1>" & vbCrLf
		Response.Write "<B>注册网址</b>" & vbCrLf
		Response.Write "</td>" & vbCrLf
		Response.Write "<td Class=TableRow1>" & vbCrLf
		Response.Write " <input type=text name='url'"
		response.write "value='"
		response.write rs("url")
		response.write "'>" & vbCrLf	
        Response.Write "</td>" & vbCrLf
		Response.Write " </tr>" & vbCrLf
		
		
		Response.Write " <tr>" & vbCrLf
		Response.Write "<td Class=TableRow1>" & vbCrLf
		Response.Write "<B>注册日期</b>" & vbCrLf
		Response.Write "</td>" & vbCrLf
		Response.Write "<td Class=TableRow1>" & vbCrLf
		Response.Write " <input type=text name='urldate'"
		response.write "value='"
		response.write rs("urldate")
		response.write "'>" & vbCrLf	
        Response.Write "</td>" & vbCrLf
		Response.Write " </tr>" & vbCrLf

		
		
		
		
		Response.Write " <tr>" & vbCrLf
		Response.Write "<td Class=TableRow1>" & vbCrLf
		Response.Write "<B>注册人</b>" & vbCrLf
		Response.Write "</td>" & vbCrLf
		Response.Write "<td Class=TableRow1>" & vbCrLf
		Response.Write " <input type=text name='urlman'"
		response.write "value='"
		response.write rs("urlman")
		response.write "'>" & vbCrLf	
        Response.Write "</td>" & vbCrLf
		Response.Write " </tr>" & vbCrLf
		
		
		
		Response.Write " <tr>" & vbCrLf
		Response.Write "<td Class=TableRow1>" & vbCrLf
		Response.Write "<B>注册模块</b>" & vbCrLf
		Response.Write "</td>" & vbCrLf
		Response.Write "<td Class=TableRow1>" & vbCrLf
		
		response.write "<input type='checkbox' name='urlflag' value='SiteConfig'"
		If InStr(urlflag, "SiteConfig") <> 0 Then 
			Response.Write " checked"
		end if
		response.write ">常规设置"
		
		

		response.write "<input type='checkbox' name='urlflag' value='yemian'"
		If InStr(urlflag, "yemian") <> 0 Then 
			Response.Write " checked"
		end if
		response.write ">单页面图文"

		response.write "<input type='checkbox' name='urlflag' value='Article'"
		If InStr(urlflag, "Article") <> 0 Then 
			Response.Write " checked"
		end if
		response.write ">新闻频道"
		
		
		response.write "<input type='checkbox' name='urlflag' value='soft'"
		If InStr(urlflag, "soft") <> 0 Then 
			Response.Write " checked"
		end if
		response.write ">下载频道"
		
		response.write "<br>"

		
		response.write "<input type='checkbox' name='urlflag' value='flash'"
		If InStr(urlflag, "flash") <> 0 Then 
			Response.Write " checked"
		end if
		response.write ">动画频道"
			
		response.write "<input type='checkbox' name='urlflag' value='shop'"
		If InStr(urlflag, "shop") <> 0 Then 
			Response.Write " checked"
		end if
		response.write ">商品频道"	
		
		response.write "<input type='checkbox' name='urlflag' value='order'"
		If InStr(urlflag, "order") <> 0 Then 
			Response.Write " checked"
		end if
		response.write ">订单频道"	


		response.write "<input type='checkbox' name='urlflag' value='job'"
		If InStr(urlflag, "job") <> 0 Then 
			Response.Write " checked"
		end if
		response.write ">招聘频道"

		
		
		
        Response.Write "</td>" & vbCrLf
		Response.Write " </tr>" & vbCrLf
		
		
		Response.Write " <tr>" & vbCrLf
		Response.Write "<td Class=TableRow1>" & vbCrLf
		Response.Write "<B>序列号</b>" & vbCrLf
		Response.Write "</td>" & vbCrLf
		Response.Write "<td Class=TableRow1>" & vbCrLf
		Response.Write " <input type=text name='urlreg' size='100'"
		response.write "value='"
		response.write rs("urlreg")
		response.write "'>" & vbCrLf	
        Response.Write "</td>" & vbCrLf
		Response.Write " </tr>" & vbCrLf


	end if
	
		
	
	Rs.Close
	Set Rs = Nothing
	
	Response.Write " <tr>" & vbCrLf
	Response.Write " <td colspan=""6"" align=center Class=TableRow1>" & vbCrLf
	Response.Write " <input type='submit' class=""button"" name=""Submit"" value=""  注 册  "" >" & vbCrLf
	Response.Write " </td>" & vbCrLf
	Response.Write " </tr>" & vbCrLf
		Response.Write " </form>" & vbCrLf
	Response.Write " </table>" & vbCrLf
End Sub


Private Sub savereg()
	Dim adminuserid
	dim zcj
	If Not ChkAdmin("9999") Then
		ErrMsg = "<li>您没有此操作权限!</li><li>如有什么问题请联系站长？</li>"
		Founderr = True
		Exit Sub
	End If

	If Request.Form("url") = "" or Request.Form("urldate") = "" or Request.Form("urlman") = "" or Request.Form("urlreg") = ""Then
		ErrMsg = "请输入相关的内容再继续！"
		Founderr = True
		Exit Sub
	Else
		'
		if not isdate(Request.Form("urldate")) then
			ErrMsg = "错误的日期数据，请输入正确的日期数据格式！"
			Founderr = True
			Exit Sub
		end if
		'
		if Request.Form("url") <>enchiasp.SiteUrl&enchiasp.InstallDir then
			ErrMsg = "错误的网站地址，请输入正确的网站地址，加HTTP！您应该注册的网站为：<br>"&enchiasp.SiteUrl&enchiasp.InstallDir
			Founderr = True
			Exit Sub
		end if 
		
		'
		zcj=md5(request.form("url")&"liuyunfan")&"-" & md5("yunliufan")&md5(request.form("urldate"))&"-"&md5("liu")&md5(request.form("urlman")) & md5("fanyun")&"-"& md5( Replace(Replace(Request("urlflag"), "'", ""), " ", ""))
		if Request.Form("urlreg") =zcj then
			
			Set Rs = Server.CreateObject("adodb.recordset")
			SQL = "SELECT * FROM ECCMS_config"
			Rs.Open SQL, conn, 1, 3
			If Not (Rs.EOF And Rs.BOF) Then
				rs("url")=request.form("url")
				rs("urldate")=request.form("urldate")
				rs("urlreg")=request.form("urlreg")
				rs("urlman")=request.form("urlman")	
				Rs("urlflag") = Replace(Replace(Request("urlflag"), "'", ""), " ", "")
				Rs.update
			End If
			Rs.Close
			Set Rs = Nothing
			Succeed ("注册成功！")
		else
			ErrMsg = "错误的序列号，请输入正确的序列号！"
			Founderr = True
			Exit Sub
		end if
	End If
	
	End Sub

%>
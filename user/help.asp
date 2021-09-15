<!--#include file="config.asp"-->
<!--#include file="check.asp"-->
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
Call InnerLocation("用户帮助")

ErrMsg = ErrMsg + "<li>对不起！您没有查看本页的权限，如有什么问题请联系管理员。</li>"
Founderr = True
If Founderr = True Then
	Call Returnerr(ErrMsg)
End If
%>
<!--#include file="foot.inc"-->
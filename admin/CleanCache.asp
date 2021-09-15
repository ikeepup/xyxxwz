<!--#include file="setup.asp"-->
<!--#include file="check.asp"-->
<LINK href="style.css" type=text/css rel=stylesheet>
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
Call RemoveAllCache()

Sub RemoveAllCache()
	Dim cachelist,i
	Call InnerHtml("UpdateInfo","<b>开始执行清理当前站点缓存</b>：")
	Cachelist=split(GetallCache(),",")
	If UBound(cachelist)>1 Then
		For i=0 to UBound(cachelist)-1
			DelCahe Cachelist(i)
			Call InnerHtml("UpdateInfo","更新 <b>"&cachelist(i)&"</b> 完成")
		Next
		Call InnerHtml("UpdateInfo","更新了"& UBound(cachelist)-1 &"个缓存对象<br>")
	Else
		Call InnerHtml("UpdateInfo","<b>当前站点全部缓存清理完成。</b>。")
	End If
End Sub

Function  GetallCache()
	Dim Cacheobj
	For Each Cacheobj in Application.Contents
		GetallCache = GetallCache & Cacheobj & ","
	Next
End Function

Sub DelCahe(MyCaheName)
	Application.Lock
	Application.Contents.Remove(MyCaheName)
	Application.unLock
End Sub

Sub InnerHtml(obj,msg)
	Response.Write "<li>"&msg&"</li>"
	Response.Flush
End Sub
%>
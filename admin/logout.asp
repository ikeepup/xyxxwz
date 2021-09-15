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
'清除COOKIES中管理员身份的验证信息.
Const Admin_Cookies_Name = "admin_enchiasp"
Session.Abandon
Session("AdminName") = ""
Session("AdminPass") = ""
Session("AdminGrade") = ""
Session("Adminflag") = ""
Session("AdminStatus") = ""
Session("AdminID") = ""
Session("AdminRandomCode") = ""
Response.Cookies(Admin_Cookies_Name) = ""
Response.Redirect ("../")
%>

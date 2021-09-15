<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title><%=enchiasp.SiteName%> - 用户管理中心</title>
<link href="style.css" type="text/css" rel="stylesheet" />
<script language="JavaScript" src="images/manage.js"></script>
<script language="JavaScript">
<!--
function preloadImg(src)
{
	var img=new Image();
	img.src=src
}
preloadImg("images/manage_arrow_left.gif");
var displayBar=true;
function switchBar(obj)
{
	if (displayBar)
	{
		parent.frame.style.display="none"
		displayBar=false;
		obj.src="images/manage_arrow_right.gif";
		obj.title="打开左边管理菜单";
	}
	else{
		parent.frame.style.display=""
		displayBar=true;
		obj.src="images/manage_arrow_left.gif";
		obj.title="关闭左边管理菜单";
	}
}
//-->
</script>
</head>
<body leftmargin="0" bottommargin="0" rightmargin="0" topmargin="0">
<table border="0" align="center" width="100%" cellspacing="0" cellpadding="0">
	<tr>		
		<td width="5%" height="22" background="images/manage_top_bg.gif"><img onclick="switchBar(this)" src="images/manage_arrow_right.gif" style="cursor:hand" title="关闭左边管理菜单" align="absMiddle"><td>
		<td width="95%" background="images/manage_top_bg.gif">当前位置：<span class="shadow" id="locationid"></span></td>
	</tr>
</table>
<table border="0" align="center" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td height="6"></td>
	</tr>
</table>
<!--#include file="setup.asp"-->
<!--#include file="check.asp"-->
<%
Admin_header
If Not ChkAdmin("999") Then
	Server.Transfer("showerr.asp")
	Response.End
End If

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
%>
<table border=0 align=center cellpadding=3 cellspacing=1 class=TableBorder>
 <th>QQ在线管理</th>  
 </table>
<%



Response.Buffer = True
Server.scriptTimeout="20"


if request.querystring("type")="manage" then
	select case request.querystring("mtype")
		case "edit" call neditinfo()
		case "add"  call addqq()
		case else call slist()
	end select
	if request.querystring("act")<>"" then

	select case request.querystring("act")
		case "editsiteinfo"call editsiteinfo()
		case "add" call addinfo()
		case "edit" call editinfo()
		case "delete" call delinfo()
	end select
end if

else
	call slist()

end if

If FoundErr = True Then
	ReturnError(ErrMsg)
End If
Admin_footer
SaveLogInfo(AdminName)
CloseConn


function leftshow(str,leftc)
	if len(str)>=leftc then
		leftshow=left(str,leftc)&".."
	else
		leftshow=str
	end if
end function

function htmlencode(reString)
	dim Str
	str=reString
	str=replace(str, "&", "&amp;")
	str=replace(str, ">", "&gt;")
	str=replace(str, "<", "&lt;")
	htmlencode=Str
end function

sub editsiteinfo()
	Dim infostrSourceFile,infoobjXML
	infostrSourceFile=Server.MapPath("xml/info.xml")
	Set infoobjXML=Server.CreateObject("Microsoft.XMLDOM")
	infoobjXML.load(infostrSourceFile)
	Dim infoobjNodes
	Set infoobjNodes=infoobjXML.selectSingleNode("xml/qqinfo/qqset[siteid ='1']")
	If Not IsNull(infoobjNodes) then
		infoobjNodes.childNodes(0).text=htmlencode(request.form("sitename"))
		infoobjNodes.childNodes(1).text=htmlencode(request.form("siteskin"))
		infoobjNodes.childNodes(2).text=htmlencode(request.form("siteshowx"))
		infoobjNodes.childNodes(3).text=htmlencode(request.form("siteshowy"))
		infoobjNodes.childNodes(4).text=htmlencode(request.form("siteww"))
		infoobjXML.save(infostrSourceFile)
		%>
		<script language="javascript">
			alert("参数修改成功！")
			location.href="admin_qq.asp?type=manage"
		</script>
		<%
	Else
		%>
		<script language="javascript">
			alert("Xml 未成功打开！")
			history.back()
		</script>
		<%
	End If
	Set infoobjNodes=nothing
	Set infoobjXML=nothing
end sub


sub addinfo()
	if trim(request.form("qq"))="" or trim(request.form("dis"))="" or trim(request.form("face"))="" then
		%>
		<script language="javascript">
			alert("没有填入必要的数据！")
			history.back()
		</script>
		<%
		response.end
	end if

	dim jtb_color
	if trim(request.form("color"))="" then
		jtb_color="#000000"
	else
		jtb_color=htmlencode(request.form("color"))
	end if

	Dim strSourceFile,objXML,oListNode,oDetailsNode,AllNodesNum
	strSourceFile=Server.MapPath("xml/qq.xml")
	Set objXML=Server.CreateObject("Microsoft.XMLDOM")
	objXML.load(strSourceFile)
	Dim objRootlist
	Set objRootlist=objXML.documentElement.selectSingleNode("qqlist")
	dim id
	If objRootlist.hasChildNodes then
		id=objRootlist.lastChild.lastChild.text+1
	Else
		id=1
	End If
	Set objRootlist=nothing
	Set oListNode=objXML.documentElement.selectSingleNode("qqlist").AppendChild(objXML.createElement("qq"))
	Set oDetailsNode=oListNode.appendChild(objXML.createElement("qid"))
	oDetailsNode.Text=htmlencode(request.form("qq"))
	Set oDetailsNode=oListNode.appendChild(objXML.createElement("dis"))
	oDetailsNode.Text=htmlencode(request.form("dis"))
	Set oDetailsNode=oListNode.appendChild(objXML.createElement("face"))
	oDetailsNode.Text=htmlencode(request.form("face"))
	Set oDetailsNode=oListNode.appendChild(objXML.createElement("color"))
	oDetailsNode.Text=jtb_color
	Set oDetailsNode=oListNode.appendChild(objXML.createElement("id"))
	oDetailsNode.Text=id
	objXML.save(strSourceFile)
	Set objRootlist=nothing
	Set oListNode=nothing
	Set oDetailsNode=nothing
	Set objXML=nothing
	%>
	<script language="javascript">
		alert("添加新的QQ号成功！")
		location.href="admin_qq.asp?type=manage"
	</script>
	<%
	response.end
end sub



sub editinfo()
	dim editid:editid=request.querystring("id")
	if not IsNumeric(editid) or editid="" then
		%>
		<script language="javascript">
			alert("非法操作！")
			history.back()
		</script>
		<%
		response.end
	else
		editid=clng(editid)
	end if
	if trim(request.form("qq"))="" or trim(request.form("dis"))="" or trim(request.form("face"))="" then
		%>
		<script language="javascript">
			alert("没有添入必要的数据！")
			history.back()
		</script>
		<%
		response.end
	end if
	Dim strSourceFile,objXML
	strSourceFile=Server.MapPath("xml/qq.xml")
	Set objXML=Server.CreateObject("Microsoft.XMLDOM")
	objXML.load(strSourceFile) 
	Dim objNodes
	Set objNodes=objXML.selectSingleNode("xml/qqlist/qq[id ='"&editid&"']")
	If Not IsNull(objNodes) then
		objNodes.childNodes(0).text=htmlencode(request.form("qq"))
		objNodes.childNodes(1).text=htmlencode(request.form("dis"))
		objNodes.childNodes(2).text=htmlencode(request.form("face"))
		objNodes.childNodes(3).text=htmlencode(request.form("color"))
		objXML.save(strSourceFile)
		Set objNodes=nothing
		Set objXML=nothing
		%>
		<script language="javascript">
			alert("修改成功！")
			location.href="admin_qq.asp?type=manage"
		</script>
		<%
	else
		%>
		<script language="javascript">
			alert("修改失败！")
			location.href="admin_qq.asp?type=manage"
		</script>
		<%
	end if
	response.end
end sub

sub delinfo()
	dim delid
	delid=request.querystring("id")
	if not IsNumeric(delid) or delid="" then
		%>
			<script language="javascript">
			alert("非法操作！")
			location.href="<%=jurl%>?type=manage"
		</script>
		<%
		response.end
	else
		delid=clng(delid)
	end if
	Dim strSourceFile,objXML
	strSourceFile=Server.MapPath("xml/qq.xml")
	Set objXML=Server.CreateObject("Microsoft.XMLDOM")
	objXML.load(strSourceFile)
	Dim objNodes
	Set objNodes=objXML.selectSingleNode("xml/qqlist/qq[id ='"&delid&"']")
	if Not IsNull(objNodes) then
		if request.querystring("yn")="" then
			%>
			<script language="javascript">
				if(confirm("确认要删除[<%=objNodes.childNodes(0).text%>]的信息吗？"))
					window.location="?type=manage&act=delete&id=<%=delid%>&yn=1"
				else
					history.back()
			</script>
			<%
		else
			objNodes.parentNode.removeChild(objNodes)
			objXML.save(strSourceFile)
			%>
			<script language="javascript">
				alert("删除成功！")
				location.href="admin_qq.asp?type=manage"
			</script>
			<%
		end if
	else
		%>
		<script language="javascript">
			alert("没有找到指定的条目！")
			location.href="admin_qq.asp?type=manage"
		</script>
		<%
	End If
	Set objNodes=nothing
	Set objXML=nothing
	response.end
end sub


sub slist()
%>
<table border=0 align=center cellpadding=3 cellspacing=1 class=TableBorder>
<tr>
  <td class=TableRow1 height="25" >调用方法：在通栏模版中增加如下语句调用
   &lt;script language="javascript" type="text/javascript" src="{$WebSiteUrl}{$InstallDir}qqonline/qq.asp"&gt;&lt;/script&gt;注意有的地方无法调用，仅用于各个栏目下使用，在根目录下可能会出现无法调用的情况。</td>

</tr>
 </table>

<table border=0 align=center cellpadding=3 cellspacing=1 class=TableBorder>
<%
Dim infostrSourceFile,infoobjXML 
infostrSourceFile=Server.MapPath("xml/info.xml")
Set infoobjXML=Server.CreateObject("Microsoft.XMLDOM")
infoobjXML.load(infostrSourceFile)
Dim infoobjNodes
Set infoobjNodes=infoobjXML.selectSingleNode("xml/qqinfo/qqset[siteid ='1']")
If Not IsNull(infoobjNodes) then
%>
    <form method="post" action="admin_qq.asp?type=manage&act=editsiteinfo">
    <table border=0 align=center cellpadding=3 cellspacing=1 class=TableBorder>    
      <tr>
        <td class=TableRow1 height="25" align=center>&nbsp;参数设置</td>
      </tr>     
      <tr>
        <td class=TableRow1 height="25">&nbsp;网站名称：<input type="text" name="sitename" size="15" value="<%=infoobjNodes.childNodes(0).text%>">&nbsp;使用皮肤：<input type="text" name="siteskin" size="5" value="<%=infoobjNodes.childNodes(1).text%>">（系统提供25个皮肤供选择，请填写数字1-25）</td>
      </tr>
      <tr>
        <td class=TableRow1 height="25">&nbsp;显示界面X坐标：<input type="text" name="siteshowx" size="10" value="<%=infoobjNodes.childNodes(2).text%>">&nbsp;Y坐标：<input type="text" name="siteshowy" size="10" value="<%=infoobjNodes.childNodes(3).text%>"></td>
      </tr>
         <td class=TableRow1 height="25">&nbsp;旺旺：<input type="text" name="siteww" size="60" value="<%=infoobjNodes.childNodes(4).text%>"><br>&nbsp;
备注：(如不使用淘宝旺旺，可将此参数留空。由于旺旺在线不支持中文,调用需要把中文进行编码才能行,所以如果你的旺旺号是中文需要到<a href="http://www.taobao.com/help/wangwang/wangwang_0628_12.php" target="_blank"><b>[旺旺在线]</b></a>编码后复制编码填上。比如旺旺号[购物系统]编码如下面的红色部分)<br>
&lt;a target="_blank" href="http://amos1.taobao.com/msg.ww?v=2&uid=<FONT color=#ff0000>%E5%95%89%A9%E7%B3%BB%E7%BB%9F</font>&s=1" >&lt;img border="0" src="http://amos1.taobao.com/online.ww?v=2&uid=<FONT color=#ff0000>%E5%95%89%A9%E7%B3%BB%E7%BB%9F</font>&s=1" alt="点击这里给我发消息" &lt;/a>
     </td>
     </tr>
      <tr>
        <td class=TableRow1 height="10"><input type="submit" name="submit" value="保存设置"><br></td>
      </tr>
    </table>
    </form>
<%
Else
%>
<script language="javascript">
alert("Xml 未成功打开！")
history.back()
</script>
<%
response.end
End If
Set infoobjNodes=nothing
Set infoobjXML=nothing
%>
    <table border=0 align=center cellpadding=3 cellspacing=1 class=TableBorder>
      <tr>
        <td class=TableRow1 height="25">&nbsp;QQ号</td>
        <td class=TableRow1 height="25">&nbsp;描述</td>
        <td class=TableRow1 height="25">&nbsp;头像</td>
        <td class=TableRow1 align="center" height="25">新增</td>
        <td class=TableRow1 align="center" height="25">编辑</td>
        <td class=TableRow1 align="center" height="25">删除</td>
      </tr>
<%
Dim strSourceFile,objXML,objRootsite,AllNodesNum
strSourceFile=Server.MapPath("xml/qq.xml")
Set objXML=Server.CreateObject("Microsoft.XMLDOM")
objXML.load(strSourceFile)
Set objRootsite=objXML.documentElement.selectSingleNode("qqlist")
AllNodesNum=objRootsite.childNodes.length-1
Dim iCount
For iCount=0 to AllNodesNum
%>
      <tr>
        <td class=TableRow1 height="25">&nbsp;<%=objRootsite.childNodes.item(iCount).childNodes.item(0).text%></td>
        <td class=TableRow1 height="25">&nbsp;<%=objRootsite.childNodes.item(iCount).childNodes.item(1).text%></td>
        <td class=TableRow1 height="25">&nbsp;<img src="images/qqface/<%=objRootsite.childNodes.item(iCount).childNodes.item(2).text%>_m.gif" border="0"></td>
         <td class=TableRow1 align="center" height="25"><a href="?type=manage&mtype=add">新增</a></td>
        <td class=TableRow1 align="center" height="25"><a href="?type=manage&mtype=edit&id=<%=objRootsite.childNodes.item(iCount).childNodes.item(4).text%>">编辑</a></td>
        <td class=TableRow1 align="center" height="25"><a href="?type=manage&act=delete&id=<%=objRootsite.childNodes.item(iCount).childNodes.item(4).text%>">删除</a></td>
      </tr>
<%
Next
Set objRootsite=nothing
Set objXML=nothing
%>      
    </table>
   
<%
end sub

sub addqq()
%>
 <form method="post" action="admin_qq.asp?type=manage&act=add">
    <table border=0 align=center cellpadding=3 cellspacing=1 class=TableBorder>
      <tr>
        <td class=TableRow1 height="10"></td>
      </tr>     
      <tr>
        <td class=TableRow1 height="25">&nbsp;添加新的QQ号</td>
      </tr>     
      <tr>
        <td class=TableRow1 height="25">&nbsp;QQ号：<input type="text" name="qq" size="20"></td>
      </tr>     
      <tr>
        <td class=TableRow1 height="25">&nbsp;描述：<input type="text" name="dis" size="25"></td>
      </tr>
      <tr>
        <td class=TableRow1 height="25">&nbsp;颜色：<input type="text" name="color" size="25"> 输入颜色代码例如：#000000</td>  
      </tr>
      <tr>
        <td class=TableRow1 height="25">&nbsp;头像：<br>   
        <%dim fcount:for fcount=1 to 100%>
        <input type="radio" name="face" value="<%=fcount%>"><img src="images/qqface/<%=fcount%>_m.gif" border="0">                                                      
        <%next%>                                                      
        </td>                                                                           
      </tr>
      <tr>
        <td class=TableRow1 height="10"></td>
      </tr>
      <tr>
        <td class=TableRow1 height="25">&nbsp;<input type="submit" name="submit" value="确定添加">&nbsp;<input type="reset" name="reset" value="取消重置"></td>
      </tr>
      </table>
    </form>
    </td>
  </tr>
</table>
<%
end sub

sub neditinfo()
%>
<table border=0 align=center cellpadding=3 cellspacing=1 class=TableBorder>
  <tr>
<%
dim neditid:neditid=request.querystring("id")
if not IsNumeric(neditid) or neditid="" then
%>
<script language="javascript">
alert("非法操作！")
history.back()
</script>
<%
response.end
else
neditid=clng(neditid)
end if
Dim strSourceFile,objXML
strSourceFile=Server.MapPath("xml/qq.xml")
Set objXML=Server.CreateObject("Microsoft.XMLDOM")
objXML.load(strSourceFile)
Dim objNodes
Set objNodes=objXML.selectSingleNode("xml/qqlist/qq[id ='"&neditid&"']")
If Not IsNull(objNodes) then
%>
    <form method="post" action="?type=manage&act=edit&id=<%=neditid%>">
    <table border=0 align=center cellpadding=3 cellspacing=1 class=TableBorder>
      <tr>
        <td class=TableRow1 height="10"></td>
      </tr>     
      <tr>
        <td class=TableRow1 height="25">&nbsp;修改指定的QQ信息</td>
      </tr>     
      <tr>
        <td class=TableRow1 height="25">&nbsp;QQ号：<input type="text" name="qq" size="20" value="<%=objNodes.childNodes(0).text%>"></td>
      </tr>     
      <tr>
        <td class=TableRow1 height="25">&nbsp;描述：<input type="text" name="dis" size="25" value="<%=objNodes.childNodes(1).text%>"> </td>                                                                         
      </tr>
      <tr>
        <td class=TableRow1 height="25">&nbsp;颜色：<input type="text" name="color" size="25" value="<%=objNodes.childNodes(3).text%>"> 输入颜色代码例如：#000000</td>  
      </tr>
      <tr>
        <td class=TableRow1 height="25">&nbsp;头像：<br>
        <%dim fcount:for fcount=1 to 100%>
        <input type="radio" name="face" value="<%=fcount%>" <%if objNodes.childNodes(2).text=cstr(fcount) then response.write" checked"%>><img src="images/qqface/<%=fcount%>_m.gif" border="0">                                      
        <%next%>                                      
        </td>                                                                           
      </tr>
      <tr>
        <td class=TableRow1 height="10"></td>
      </tr>
      <tr>
        <td class=TableRow1 height="25">&nbsp;<input type="submit" name="submit" value="确定修改">&nbsp;<input type="reset" name="reset" value="取消重置"></td>
      </tr>
       </table>
    </form>
    </td>
  </tr>
</table>
<%
else
%>
<script language="javascript">
alert("发生错误！")
history.back()
</script>
<%
end if
	Set objNodes=nothing
	Set objXML=nothing
end sub

%>
</body>
</html>

























































































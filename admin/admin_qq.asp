<!--#include file="setup.asp"-->
<!--#include file="check.asp"-->
<%
Admin_header
If Not ChkAdmin("999") Then
	Server.Transfer("showerr.asp")
	Response.End
End If

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
%>
<table border=0 align=center cellpadding=3 cellspacing=1 class=TableBorder>
 <th>QQ���߹���</th>  
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
			alert("�����޸ĳɹ���")
			location.href="admin_qq.asp?type=manage"
		</script>
		<%
	Else
		%>
		<script language="javascript">
			alert("Xml δ�ɹ��򿪣�")
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
			alert("û�������Ҫ�����ݣ�")
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
		alert("����µ�QQ�ųɹ���")
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
			alert("�Ƿ�������")
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
			alert("û�������Ҫ�����ݣ�")
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
			alert("�޸ĳɹ���")
			location.href="admin_qq.asp?type=manage"
		</script>
		<%
	else
		%>
		<script language="javascript">
			alert("�޸�ʧ�ܣ�")
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
			alert("�Ƿ�������")
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
				if(confirm("ȷ��Ҫɾ��[<%=objNodes.childNodes(0).text%>]����Ϣ��"))
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
				alert("ɾ���ɹ���")
				location.href="admin_qq.asp?type=manage"
			</script>
			<%
		end if
	else
		%>
		<script language="javascript">
			alert("û���ҵ�ָ������Ŀ��")
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
  <td class=TableRow1 height="25" >���÷�������ͨ��ģ������������������
   &lt;script language="javascript" type="text/javascript" src="{$WebSiteUrl}{$InstallDir}qqonline/qq.asp"&gt;&lt;/script&gt;ע���еĵط��޷����ã������ڸ�����Ŀ��ʹ�ã��ڸ�Ŀ¼�¿��ܻ�����޷����õ������</td>

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
        <td class=TableRow1 height="25" align=center>&nbsp;��������</td>
      </tr>     
      <tr>
        <td class=TableRow1 height="25">&nbsp;��վ���ƣ�<input type="text" name="sitename" size="15" value="<%=infoobjNodes.childNodes(0).text%>">&nbsp;ʹ��Ƥ����<input type="text" name="siteskin" size="5" value="<%=infoobjNodes.childNodes(1).text%>">��ϵͳ�ṩ25��Ƥ����ѡ������д����1-25��</td>
      </tr>
      <tr>
        <td class=TableRow1 height="25">&nbsp;��ʾ����X���꣺<input type="text" name="siteshowx" size="10" value="<%=infoobjNodes.childNodes(2).text%>">&nbsp;Y���꣺<input type="text" name="siteshowy" size="10" value="<%=infoobjNodes.childNodes(3).text%>"></td>
      </tr>
         <td class=TableRow1 height="25">&nbsp;������<input type="text" name="siteww" size="60" value="<%=infoobjNodes.childNodes(4).text%>"><br>&nbsp;
��ע��(�粻ʹ���Ա��������ɽ��˲������ա������������߲�֧������,������Ҫ�����Ľ��б��������,������������������������Ҫ��<a href="http://www.taobao.com/help/wangwang/wangwang_0628_12.php" target="_blank"><b>[��������]</b></a>������Ʊ������ϡ�����������[����ϵͳ]����������ĺ�ɫ����)<br>
&lt;a target="_blank" href="http://amos1.taobao.com/msg.ww?v=2&uid=<FONT color=#ff0000>%E5%95%89%A9%E7%B3%BB%E7%BB%9F</font>&s=1" >&lt;img border="0" src="http://amos1.taobao.com/online.ww?v=2&uid=<FONT color=#ff0000>%E5%95%89%A9%E7%B3%BB%E7%BB%9F</font>&s=1" alt="���������ҷ���Ϣ" &lt;/a>
     </td>
     </tr>
      <tr>
        <td class=TableRow1 height="10"><input type="submit" name="submit" value="��������"><br></td>
      </tr>
    </table>
    </form>
<%
Else
%>
<script language="javascript">
alert("Xml δ�ɹ��򿪣�")
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
        <td class=TableRow1 height="25">&nbsp;QQ��</td>
        <td class=TableRow1 height="25">&nbsp;����</td>
        <td class=TableRow1 height="25">&nbsp;ͷ��</td>
        <td class=TableRow1 align="center" height="25">����</td>
        <td class=TableRow1 align="center" height="25">�༭</td>
        <td class=TableRow1 align="center" height="25">ɾ��</td>
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
         <td class=TableRow1 align="center" height="25"><a href="?type=manage&mtype=add">����</a></td>
        <td class=TableRow1 align="center" height="25"><a href="?type=manage&mtype=edit&id=<%=objRootsite.childNodes.item(iCount).childNodes.item(4).text%>">�༭</a></td>
        <td class=TableRow1 align="center" height="25"><a href="?type=manage&act=delete&id=<%=objRootsite.childNodes.item(iCount).childNodes.item(4).text%>">ɾ��</a></td>
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
        <td class=TableRow1 height="25">&nbsp;����µ�QQ��</td>
      </tr>     
      <tr>
        <td class=TableRow1 height="25">&nbsp;QQ�ţ�<input type="text" name="qq" size="20"></td>
      </tr>     
      <tr>
        <td class=TableRow1 height="25">&nbsp;������<input type="text" name="dis" size="25"></td>
      </tr>
      <tr>
        <td class=TableRow1 height="25">&nbsp;��ɫ��<input type="text" name="color" size="25"> ������ɫ�������磺#000000</td>  
      </tr>
      <tr>
        <td class=TableRow1 height="25">&nbsp;ͷ��<br>   
        <%dim fcount:for fcount=1 to 100%>
        <input type="radio" name="face" value="<%=fcount%>"><img src="images/qqface/<%=fcount%>_m.gif" border="0">                                                      
        <%next%>                                                      
        </td>                                                                           
      </tr>
      <tr>
        <td class=TableRow1 height="10"></td>
      </tr>
      <tr>
        <td class=TableRow1 height="25">&nbsp;<input type="submit" name="submit" value="ȷ�����">&nbsp;<input type="reset" name="reset" value="ȡ������"></td>
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
alert("�Ƿ�������")
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
        <td class=TableRow1 height="25">&nbsp;�޸�ָ����QQ��Ϣ</td>
      </tr>     
      <tr>
        <td class=TableRow1 height="25">&nbsp;QQ�ţ�<input type="text" name="qq" size="20" value="<%=objNodes.childNodes(0).text%>"></td>
      </tr>     
      <tr>
        <td class=TableRow1 height="25">&nbsp;������<input type="text" name="dis" size="25" value="<%=objNodes.childNodes(1).text%>"> </td>                                                                         
      </tr>
      <tr>
        <td class=TableRow1 height="25">&nbsp;��ɫ��<input type="text" name="color" size="25" value="<%=objNodes.childNodes(3).text%>"> ������ɫ�������磺#000000</td>  
      </tr>
      <tr>
        <td class=TableRow1 height="25">&nbsp;ͷ��<br>
        <%dim fcount:for fcount=1 to 100%>
        <input type="radio" name="face" value="<%=fcount%>" <%if objNodes.childNodes(2).text=cstr(fcount) then response.write" checked"%>><img src="images/qqface/<%=fcount%>_m.gif" border="0">                                      
        <%next%>                                      
        </td>                                                                           
      </tr>
      <tr>
        <td class=TableRow1 height="10"></td>
      </tr>
      <tr>
        <td class=TableRow1 height="25">&nbsp;<input type="submit" name="submit" value="ȷ���޸�">&nbsp;<input type="reset" name="reset" value="ȡ������"></td>
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
alert("��������")
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

























































































<!--#include file="setup.asp"-->
<!--#include file="check.asp"-->
<%
Admin_header
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

Dim maxperpage,CurrentPage,totalnumber,TotalPageNum
dim i,strClass,action
Action = LCase(Request("action"))


'Ȩ���ж�
If not ChkAdmin("fengmian") Then
	Server.Transfer("showerr.asp")
	Response.End
End If

'	| <a href='?action=fengmiancanshu'><font color=blue>�����������</font></a>    
%>

<table border=0 align=center cellpadding=3 cellspacing=1 class=tableborder>
<tr>
	<th>����ģ�����</th>
</tr>
<tr>
	<td class=tablerow2><strong>��������</strong> <a href='?'>������ҳ</a>                                                                                                                                                                                         
	| <a href='?action=fengmiankaiguan'><font color=blue>���濪������</font></a>                                                                                                                                                                                                                                                                                                                                                          
	| <a href='?action=fengmianmoban'><font color=blue>ѡ�����ģ��</font></a>                                          
	| <a href='?action=addfengmianmoban'><font color=blue>�½�����ģ��</font></a>                                     
	| <a href='Admin_UploadFile.Asp?ChannelID=-1&UploadDir=UploadPic'><font color=blue>�ϴ�ͼƬ����</font></a>                                                                                                                                                                                             
	</td>
</tr>
</table>
<br>
<%
Select Case Trim(Action)
	case "fengmiankaiguan"
		call fengmiankaiguan()
	case "setfengmiankaiguan"
		call setfengmiankaiguan()
	case "fengmianmoban"
		call fengmianmoban()
	case "fengmiancanshu"
		call fengmiancanshu()
	case "setfengmiancanshu"
		call setfengmiancanshu()
	case "setfengmianmoban"
		call setfengmianmoban()
	case "editfengmianmoban"
		call editfengmianmoban()
	case "saveeditfengmiancanshu"
		call saveeditfengmiancanshu()
	case "huanyuanfengmianmoban"
		call huanyuanfengmianmoban()
	case "addfengmianmoban"
		call addfengmianmoban()
	case "addfengmian"
		call addfengmian()
	case "editaddfengmianmoban"
		call editaddfengmianmoban()
	case "delfengmianmoban"
		call delfengmianmoban()
	Case Else
		Call MainPage
End Select
If FoundErr = True Then
	ReturnError(ErrMsg)
End If
Admin_footer
SaveLogInfo(AdminName)
CloseConn
Private Sub MainPage()
%>
<table border=0 align=center cellpadding=3 cellspacing=1 class=tableborder>
<td class=tablerow2>
<p>������������ </p>
<p>��ҳ���������÷��棬���������Ƿ����÷��棬ѡ�����ģ����Զ���ģ�壬�趨���汳����ɫ������ͼƬ���������ֵ�</p>
<p>1���������÷��濪��  </p>
<p>2������ѡ��ϵͳĬ�ϵ�ģ��  </p>
<p>3��Ҳ�����Լ�����һ��ģ��  </p>
</td>
</table>


<%       
end sub
%>






<%       
Private Sub fengmiankaiguan()
%>
<table border=0 align=center cellpadding=3 cellspacing=1 class=tableborder>
  <tbody>
    <tr class="title">
      <th>���濪������</th>
    </tr>
    <tr>
      <td colSpan="2">
        <table >
          <form name="cn" action="?action=setfengmiankaiguan" method="post">
            <tbody>
              <tr class="tdbg">
                <td width="190" class=tablerow2><strong>�Ƿ����÷���ҳ�棺</strong></td>
                <td width="598" class=tablerow2>
                
            <input type="radio" name="usefengmian" value="0" <%If enchiasp.usefengmian = "0" Then Response.Write (" checked")%>> �ر� 
			<input type="radio" name="usefengmian" value="1" <%If enchiasp.usefengmian = "1" Then Response.Write (" checked")%>> ��                 
                </td>
              </tr>      
              <br>       
                 <tr class="tdbg">
                <td align="middle" width="190" class=tablerow2>��</td>
                <td width="598" class=tablerow2><input type="submit" value="�ύ" name="Submit">��                                                                                                                                                                                                                                         
                  <input type="reset" value="����" name="Submit"></td>
              </tr>
              </FORM>
            </tbody>
          </table>
        </td>
      </tr>
    </tbody>
  </table>
<% 
end sub
private sub setfengmiankaiguan()
	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL = "select * from [ECCMS_Config] where id = 1"
	Rs.Open SQL,Conn,1,3
	Rs("usefengmian") = Trim(Request.Form("usefengmian"))
	Rs.update
	Rs.close:set Rs = Nothing
	Application.Contents.RemoveAll
	Succeed("<li>��ϲ�����������óɹ���</li>")

end sub
%>
  
<%       
Private Sub fengmiancanshu()
%>

<table border=0 align=center cellpadding=3 cellspacing=1 class=tableborder>
  <tbody>
    <tr class="title">
      <th>�����������</th>
    </tr>
    <tr>
      <td colSpan="2">
        <table >
          <form name="cn" action="?action=setfengmiancanshu" method="post">
            <tbody>              
              <tr class="tdbg">
                <td width="190" class=tablerow2><strong>������������</strong></td>
                <td width="598" class=tablerow2><input name="fengmianname" type="text" id="fengmianname" size="50" value='<%=enchiasp.fengmianname%>'></td>
              </tr>
              <tr class="tdbg">
                <td width="190" class=tablerow2><strong>������λ�ã�</strong></td>
                <td width="598" class=tablerow2>�붥�߾ࣺ <input  id="fengmiannametop" size="6" value="<%=enchiasp.fengmiannametop%>" name="fengmiannametop">                                                                                                                                                                                                                                         
                  PX���� ����߾ࣺ <input  id="fengmiannameleft" size="6" value="<%=enchiasp.fengmiannameleft%>" name="fengmiannameleft">                                                                                                                                                                                                                                         
                  PX</td>                                                                                                                                                                                                                                        
              </tr>
              <tr class="tdbg">
                <td align="middle" width="190" class=tablerow2>��</td>
                <td width="598" class=tablerow2><input type="submit" value="�ύ" name="Submit">��                                                                                                                                                                                                                                         
                  <input type="reset" value="����" name="Submit"></td>
              </tr>
              </FORM>
            </tbody>
          </table>
        </td>
      </tr>
    </tbody>
  </table>
<% 

end sub
private sub setfengmiancanshu()
	If Not (IsNumeric(Trim(Request.Form("fengmiannametop"))) and IsNumeric(Trim(Request.Form("fengmiannameleft")))) Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>�����ϵͳ����!������������</li>"
	else
		Set Rs = Server.CreateObject("ADODB.Recordset")
		SQL = "select * from [ECCMS_Config] where id = 1"
		Rs.Open SQL,Conn,1,3
		Rs("fengmianname") = Trim(Request.Form("fengmianname"))
		Rs("fengmiannametop") = Trim(Request.Form("fengmiannametop"))
		Rs("fengmiannameleft") = Trim(Request.Form("fengmiannameleft"))
		Rs.update
		Rs.close:set Rs = Nothing
		Application.Contents.RemoveAll
		Succeed("<li>��ϲ�����������óɹ���</li>")
	End If

	

end sub

%>
  







 <% 
Private Sub fengmianmoban()
	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL = "select * from [ECCMS_fengmian] where isuse=1"
	If IsSqlDataBase = 1 Then
		If CurrentPage > 100 Then
			Rs.Open SQL, Conn, 1, 1
		Else
			Set Rs = Conn.Execute(SQL)
		End If
	Else
		Rs.Open SQL, Conn, 1, 1
	End If
	If Rs.bof or Rs.EOF Then
		Response.Write "<table border=0 align=center cellpadding=3 cellspacing=1 class=tableborder><tr><td align=center colspan=5 class=TableRow2>��û��ѡ�����ģ�壡</td></tr></table>"
	else
		Response.Write "<table border=0 align=center cellpadding=3 cellspacing=1 class=tableborder><tr><td align=center colspan=5 class=TableRow2>��Ŀǰѡ��ķ���ģ��Ϊ��(<font color=red>"& rs("name") &"</font>),ģ����Ϊ��<font color=red>"& rs("bh") &"</font></td></tr></table>"

	end if
	Rs.close:set Rs = Nothing

%>
  
  <table border=0 align=center cellpadding=3 cellspacing=1 class=tableborder>
<tr>
	<th class=tablerow1>ģ����</th>
	<th class=tablerow1>ģ�����</th>
	<th class=tablerow1>ģ����ͼ</th>
	<th class=tablerow1>����ѡ��</th>

</tr>
<%
	maxperpage = 5 '###ÿҳ��ʾ��
	
	If Not IsNumeric(Request("page")) And Len(Request("page")) <> 0 Then
		Response.Write ("�����ϵͳ����!����������")
		Response.End
	End If
	If Not IsEmpty(Request("page")) And Len(Request("page")) <> 0 Then
		CurrentPage = CInt(Request("page"))
	Else
		CurrentPage = 1
	End If
	If CInt(CurrentPage) = 0 Then CurrentPage = 1
	TotalNumber = enchiasp.Execute("Select Count(id) from ECCMS_fengmian")(0)
	TotalPageNum = CInt(TotalNumber / maxperpage)  '�õ���ҳ��
	If TotalPageNum < TotalNumber / maxperpage Then TotalPageNum = TotalPageNum + 1
	If CurrentPage < 1 Then CurrentPage = 1
	If CurrentPage > TotalPageNum Then CurrentPage = TotalPageNum
	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL = "select * from [ECCMS_fengmian] order by id desc"
	If IsSqlDataBase = 1 Then
		If CurrentPage > 100 Then
			Rs.Open SQL, Conn, 1, 1
		Else
			Set Rs = Conn.Execute(SQL)
		End If
	Else
		Rs.Open SQL, Conn, 1, 1
	End If
	enchiasp.SqlQueryNum = enchiasp.SqlQueryNum + 1
	If Rs.bof or Rs.EOF Then
		Response.Write "<tr><td align=center colspan=5 class=TableRow2>��û�з���ģ�壡</td></tr>"
	Else
		If TotalPageNum > 1 then Rs.Move (CurrentPage - 1) * maxperpage
		i = 0
%>
<tr>
	<td colspan=5 class=tablerow2><%Call showpage()%></td>
</tr>
<%
	Do While Not Rs.EOF And i < CInt(maxperpage)
		If Not Response.IsClientConnected Then Response.End
		If (i mod 2) = 0 Then
			strClass = "class=TableRow1"
		Else
			strClass = "class=TableRow2"
		End If
		Response.Write "<tr>"
		
		Response.Write "	<td align=center " & strClass & "><a href='?action=setfengmianmoban&id="&rs("id")&""
		Response.Write "	'>"
		Response.Write Rs("bh")
		Response.Write "	</a></td>"
		
		
		Response.Write "	<td align=center " & strClass & "><a href='?action=setfengmianmoban&id="&rs("id")&""
		Response.Write "	'>"
		Response.Write Rs("name")
		if rs("isuse")="1" then
			Response.Write  "<br><font color=red>"
			Response.Write "(��ǰѡ��ģ��)"
			Response.Write "</font>"
		end if
		Response.Write "	</a></td>"
		if  not Rs("issystem")="1" then
		%>
		<td align="center" <%=strClass%>> <% if Rs("slt")<>"" then %><a href=<%=enchiasp.ChannelPath%><%= Rs("slt")%> target=_blank><img src=<%=enchiasp.ChannelPath%><%= Rs("slt")%> width="200" height="150" align="center" border="1"></a><% else response.write "û������ͼ" end if%></tD>
	<%
	else
	%>
		<td align="center" <%=strClass%>> <% if Rs("slt")<>"" then %><a href=<%=enchiasp.installdir%><%= Rs("slt")%> target=_blank><img src=<%=enchiasp.installdir%><%= Rs("slt")%> width="200" height="150" align="center" border="1"></a><% else response.write "û������ͼ" end if%></tD>
	
	<%
	end if
	%>
	<td align=center <%=strClass%>><% if Rs("issystem")="1" then %><a href='?action=editfengmianmoban&id=<%=Rs("id")%>'> <% else %></a><a href='?action=editaddfengmianmoban&id=<%=Rs("id")%>'><%end if%> �༭</a>
	 <% if Rs("issystem")="1" then %>|<a href='?action=huanyuanfengmianmoban&id=<%=Rs("id")%>'>��ԭ</a> <% else %> |<a href='?action=delfengmianmoban&id=<%=Rs("id")%>' onclick="{if(confirm('ɾ���󽫲��ָܻ�����ȷ��Ҫɾ����ģ����?')){return true;}return false;}">ɾ��</a><%end if%>                                                          
	</td>      
	
<%
		Rs.movenext
			i = i + 1
			If i >= maxperpage Then Exit Do
		Loop
	End If
	Rs.Close:Set Rs = Nothing
%>
<tr>
	<td colspan=5 class=tablerow2><%Call showpage()%></td>
</tr>
</table>
  
<% 
end sub

private sub setfengmianmoban
	if LCase(Request("id"))<>"" then
	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL = "select * from [ECCMS_fengmian]"
	Rs.Open SQL,Conn,1,3
	do until rs.eof
		Rs("isuse") =0
		Rs.update
		rs.movenext
	loop
	
	rs.close
	SQL = "select * from [ECCMS_fengmian] where id ="&LCase(Request("id"))&""
	Rs.Open SQL,Conn,1,3
	Rs("isuse") =1
	Rs.update
	Rs.close:set Rs = Nothing
	Application.Contents.RemoveAll
	Succeed("<li>��ϲ����ģ��ѡ��ɹ���</li>")
	end if

end sub

private sub huanyuanfengmianmoban()
dim rst,i
dim shuju(29)
if LCase(Request("id"))<>"" then
	Set Rst = Server.CreateObject("ADODB.Recordset")
	SQL = "select * from [ECCMS_fengmianmoren] where id ="&LCase(Request("id"))&""
	Rst.Open SQL,Conn,1,1
	if rst.eof or rst.bof then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>�Բ��𣬲��ܽ��л�ԭ����ģ���ϵͳĬ��ģ�壡</li>"
	else
		for i=0 to 28
			shuju(i)=rst(i)
		next
		Set Rs = Server.CreateObject("ADODB.Recordset")
		SQL = "select * from [ECCMS_fengmian] where id ="&LCase(Request("id"))&""
		Rs.Open SQL,Conn,1,3
		if rs.eof or rs.bof then
			FoundErr = True
			ErrMsg = ErrMsg + "<li>�Բ��𣬲��ܽ��л�ԭ����ģ���ϵͳĬ��ģ�壡</li>"
		else
			for i=1 to 28
				if i<>5 then
					rs(i)=shuju(i)
				end if
			next
			Rs.update
			Application.Contents.RemoveAll
			Succeed("<li>��ϲ����ģ�廹ԭ�ɹ���</li>")
		end if	
		Rs.close:set Rs = Nothing
	end if

	Rst.close:set Rst = Nothing
	
end if

end sub


private sub editfengmianmoban
	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL = "select * from [ECCMS_fengmian] where id ="&LCase(Request("id"))&""
	If IsSqlDataBase = 1 Then
		If CurrentPage > 100 Then
			Rs.Open SQL, Conn, 1, 1
		Else
			Set Rs = Conn.Execute(SQL)
		End If
	Else
		Rs.Open SQL, Conn, 1, 1
	End If
	If Rs.bof or Rs.EOF Then
		Response.Write "<table border=0 align=center cellpadding=3 cellspacing=1 class=tableborder><tr><td align=center colspan=5 class=TableRow2>��û��ѡ�����ģ�壡</td></tr></table>"
	else
		Response.Write "<table border=0 align=center cellpadding=3 cellspacing=1 class=tableborder><tr><td align=center colspan=5 class=TableRow2>��Ŀǰ�����޸ĵķ���ģ��Ϊ��(<font color=red>"& rs("name") &"</font>),ģ����Ϊ��<font color=red>"& rs("bh") &"</font></td></tr></table>"

	end if
	

%>
<script language= JavaScript>
function SelectPhoto(){
  var arr=showModalDialog('Admin_selFile.asp?ChannelID=-1&UploadDir=UploadPic', '', 'dialogWidth:800px; dialogHeight:600px; help: no; scroll: yes; status: no');
  if(arr!=null){
    var ss=arr.split('|');
    document.myform.ImageUrl.value=ss[0];
    }
}
</script>

<table border=0 align=center cellpadding=3 cellspacing=1 class=tableborder>
  <tbody>
    <tr class="title">
      <th>����ģ���޸�(�����޸���Ҫ�߱�һ������ҳ֪ʶ���������޸�)</th>
    </tr>
    
    <table border=0 align=center cellpadding=3 cellspacing=1 class=tableborder>
     <tr>
      <td>�����ǩ����</td>
     </tr>
     <tr>
      <td class=tablerow2>{$fengmianinstalldir}</td>
      <td class=tablerow2>����ģ������·��</td>
      <td class=tablerow2>{$Copyright}</td>
      <td class=tablerow2>��Ȩ��Ϣ</td>
    </tr>
	
	 <tr>
      <td class=tablerow2>{$InstallDir}</td>
      <td class=tablerow2>��վ����·��</td>
      <td class=tablerow2>{$fengmiancss}</td>
      <td class=tablerow2>��ʽ��Ϣ</td>
    </tr>

	<tr>
      <td class=tablerow2>{$fengmianbg}</td>
      <td class=tablerow2>ģ�屳����ɫ</td>
      <td class=tablerow2>{$WebSiteName}</td>
      <td class=tablerow2>��վ����</td>
    </tr>
	
	<tr>
      <td class=tablerow2>{$fengmianpic1}</td>
      <td class=tablerow2>��1��ͼƬ</td>
      <td class=tablerow2>{$fengmianflash1}</td>
      <td class=tablerow2>��1��FLASH</td>
    </tr>
	<tr>
    <td class=tablerow2>{$fengmianbgimg}</td>
      <td class=tablerow2>����ͼƬ</td>
      <td class=tablerow2>{$fengmianlogo}</td>
      <td class=tablerow2>��վLOGO</td>
    </tr>
	<tr>
    <td class=tablerow2>{$fengmianbgmidi}</td>
      <td class=tablerow2>��������</td>
      <td class=tablerow2></td>
      <td class=tablerow2></td>
    </tr>

	<tr>
    <td class=tablerow2>     &lt;script src="{$InstallDir}inc/channel.js" type="text/javascript">&lt;/script></td>                    
      <td class=tablerow2>Ƶ���б�����</td>
      <td class=tablerow2>&lt;script src="{$InstallDir}count.asp" type="text/javascript">&lt;/script></td>                        
      <td class=tablerow2>������</td>
    </tr>





	</table>
	<br>
	
    <tr>
      <td colSpan="2">
        <table border=0 align=center cellpadding=3 cellspacing=1 class=tableborder >
          <form name="myform" action="?action=saveeditfengmiancanshu&id=<%=rs("id")%>" method="post">
            <tbody>   
                       
              <tr >
                <td width="190" class=tablerow2><strong>ģ������</strong></td>
                <td width="598" class=tablerow2>
                <input name="name" type="text" id="name" size="50" value='<%=rs("name")%>'>
                </td>
              </tr>
              
              <tr class="tdbg" >
                <td width="190" class=tablerow2><strong>ģ���ţ�</strong></td>
                <td width="598" class=tablerow2>                                                              
                    <input name="bh" type="text" id="bh" size="50" value='<%=rs("bh")%>'>
                </td>                                                                                                              
              </tr>
              <tr class="tdbg">
                <td width="190" class=tablerow2><strong>����ͼ��</strong></td>
                <td width="598" class=tablerow2>                                                              
                    <input name="imageurl" type="text" id="ImageUrl" size="50" value='<%=rs("slt")%>'>
                    <input type='button' name='selectpic' value='�����ϴ�ͼƬ��ѡ��' onclick='SelectPhoto()' class=button>
                    <%
                    if rs("issystem")=1 then
                     	if Rs("slt")<>"" then 
                     		%>
                     		<img src=<%=enchiasp.InstallDir%><%=Rs("slt")%> width="100" height="75" align="center" border="1">  <br>
                     		<% 
                     	end if
                     else
                          if Rs("slt")<>"" then 
                          	%>
                          	
                          	<img src=<%=Rs("slt")%> width="100" height="75" align="center" border="1"><br> 
                          	<% 
                          end if
                     end if
                    %>
                    <iframe name="image" frameborder=0 width='100%' height=45 scrolling=no src=upload.asp?ChannelID=-1></iframe>
                </td>                                                                                                              
              </tr>

                <tr class="tdbg" >
                <td width="190" class=tablerow2><strong>����Ŀ¼��</strong></td>
                <td width="598" class=tablerow2>                                                              
                    <input name="usedir" type="text" id="usedir" size="50" value='<%=rs("usedir")%>'>
                </td>                                                                                                              
              </tr>
              <%
              dim i
              for i=1 to 10
				if rs("pic"&i&"")<>"" then
					response.write "<tr class='tdbg'>"
					 response.write "<td width='190' class=tablerow2><strong>ͼƬ"& i&"��</strong></td>"
					 response.write "<td width='598' class=tablerow2>     "
					 response.write " <input type='text' name='pic"
					 response.write i
					 response.write "' size='50' value='"
					 response.write  rs("pic"& i &"")
					 response.write "'>"
					 response.write "</td></tr>"
				end if
				next 
			
			for i=1 to 5
				if rs("flash"&i&"")<>"" then
					response.write "<tr class='tdbg'>"
					 response.write "<td width='190' class=tablerow2><strong>FLASH"& i&"��</strong></td>"
					 response.write "<td width='598' class=tablerow2>     "
					 response.write " <input  type='text' name='flash"
					 response.write i
					 response.write "' size='50' value='"
					 response.write  rs("flash"& i &"")
					 response.write "'>"
					 response.write "</td></tr>"
				end if
			next 


              %>
              	<tr class="tdbg">
                <td width="190" class=tablerow2><strong>������ɫ��</strong></td>
                <td width="598" class=tablerow2>                                                              
                    <input name="bg" type="text" id="bg" size="50" value='<%=rs("bg")%>'>
                </td>                                                                                                              
              </tr>
              <tr class="tdbg">
                <td width="190" class=tablerow2><strong>����ͼƬ��</strong></td>
                <td width="598" class=tablerow2>                                                              
                    <input name="bgimg" type="text" id="bgimg" size="50" value='<%=rs("bgimg")%>'>
                </td>                                                                                                              
              </tr>
               <tr class="tdbg">
                <td width="190" class=tablerow2><strong>�������֣�</strong></td>
                <td width="598" class=tablerow2>                                                              
                    <input name="bgmidi" type="text" id="bgmidi" size="50" value='<%=rs("bgmidi")%>'>
                </td>                                                                                                              
              </tr>

              <tr class="tdbg">
                <td width="190" class=tablerow2><strong>��վLOGO��</strong></td>
                <td width="598" class=tablerow2>                                                              
                    <input name="logo" type="text" id="logo" size="50" value='<%=rs("logo")%>'>
                </td>                                                                                                              
              </tr>


              <tr class="tdbg">
                <td width="190" class=tablerow2><strong>��ʽ��</strong></td>
                <td width="598" class=tablerow2>  
                  <textarea rows="10" name="css" cols="90"><%=rs("css")%></textarea>                                                                
                </td>                                                                                                              
              </tr>
              
 			

              <tr class="tdbg">
                <td width="190" class=tablerow2><strong>ģ�����ݣ�</strong></td>
                <td width="598" class=tablerow2>  
                <textarea rows="10" name="nr" cols="90"><%=rs("nr")%></textarea>                                                            
                </td>                                                                                                              
              </tr>
              
			<tr class="tdbg">
                <td width="190" class=tablerow2><strong>��ע��</strong></td>
                <td width="598" class=tablerow2>  
                <textarea rows="3" name="bz" cols="90"><%=rs("bz")%></textarea>                                                            
                </td>                                                                                                              
              </tr>

              
              
              <tr class="tdbg">
                <td align="middle" width="190" class=tablerow2>��</td>
                <td width="598" class=tablerow2><input type="submit" value="�ύ" name="Submit">��                                                                                                                                                                                                                                         
                  <input type="reset" value="����" name="Submit"></td>
              </tr>
              </FORM>
            </tbody>
          </table>
        </td>
      </tr>
    </tbody>
  </table>



<%
Rs.close:set Rs = Nothing
end sub
Private Sub saveeditfengmiancanshu()
		dim i
		if Trim(Request.Form("name"))="" or Trim(Request.Form("bh"))="" or Trim(Request.Form("usedir"))="" or Trim(Request.Form("nr"))="" then
			FoundErr = True
			ErrMsg = ErrMsg + "<li>����Ĳ�������!���������ݣ���Щ���ݲ���Ϊ�գ�</li>"

		else
		Set Rs = Server.CreateObject("ADODB.Recordset")
		SQL = "select * from [ECCMS_fengmian] where id ="&LCase(Request("id"))&""
		Rs.Open SQL,Conn,1,3
	
		rs("name")=Trim(Request.Form("name"))
		
		rs("bh")=Trim(Request.Form("bh"))
		
		if Trim(Request.Form("ImageUrl"))="" then
			rs("slt")=null
		else

			rs("slt")=Trim(Request.Form("ImageUrl"))
		end if
		rs("usedir")=Trim(Request.Form("usedir"))
		if Trim(Request.Form("bg"))="" then
		rs("bg")=null
		else
		rs("bg")=Trim(Request.Form("bg"))
		end if
		rs("css")=Trim(Request.Form("css"))

		rs("nr")=Trim(Request.Form("nr"))
		if Trim(Request.Form("bz"))="" then
		rs("bz")=null
		else

		rs("bz")=Trim(Request.Form("bz"))
		end if
		for i=1 to 10
			if rs("pic"&i&"")<>"" then
				 rs("pic"&i&"") = Trim(Request.Form("pic"& i))

			end if
		next 
		
		for i=1 to 5
			if rs("flash"&i&"")<>"" then
				 rs("flash"&i&"") = Trim(Request.Form("flash"& i))
			end if
		next 

		if Trim(Request.Form("bgimg"))="" then
		rs("bgimg")=null
		else

		rs("bgimg")=Trim(Request.Form("bimg"))
		end if
		
		if Trim(Request.Form("logo"))="" then
		rs("logo")=null
		else

		rs("logo")=Trim(Request.Form("logo"))
		end if
		
		if Trim(Request.Form("bgmidi"))="" then
		rs("bgmidi")=null
		else

		rs("bgmidi")=Trim(Request.Form("bgmidi"))
		end if

		

		Rs.update
		Rs.close:set Rs = Nothing
		Application.Contents.RemoveAll
		Succeed("<li>��ϲ�����������óɹ���</li>")
		end if

end sub
private sub addfengmianmoban()
%>
<script language= JavaScript>
function SelectPhoto(){
  var arr=showModalDialog('Admin_selFile.asp?ChannelID=-1&UploadDir=UploadPic', '', 'dialogWidth:800px; dialogHeight:600px; help: no; scroll: yes; status: no');
  if(arr!=null){
    var ss=arr.split('|');
    document.myform.ImageUrl.value=ss[0];
  }
}
</script>
<table border=0 align=center cellpadding=3 cellspacing=1 class=tableborder>
     <tr>
      <td>�����ǩ����</td>
     </tr>
     <tr>
      <td class=tablerow2>{$fengmianinstalldir}</td>
      <td class=tablerow2>����ģ������·��</td>
      <td class=tablerow2>{$Copyright}</td>
      <td class=tablerow2>��Ȩ��Ϣ</td>
    </tr>
	
	 <tr>
      <td class=tablerow2>{$InstallDir}</td>
      <td class=tablerow2>��վ����·��</td>
      <td class=tablerow2>{$fengmiancss}</td>
      <td class=tablerow2>��ʽ��Ϣ</td>
    </tr>

	<tr>
      <td class=tablerow2>{$fengmianbg}</td>
      <td class=tablerow2>ģ�屳����ɫ</td>
      <td class=tablerow2>{$WebSiteName}</td>
      <td class=tablerow2>��վ����</td>
    </tr>
	
	<tr>
      <td class=tablerow2>{$fengmianpic1}</td>
      <td class=tablerow2>��1��ͼƬ</td>
      <td class=tablerow2>{$fengmianflash1}</td>
      <td class=tablerow2>��1��FLASH</td>
    </tr>
	<tr>
    <td class=tablerow2>{$fengmianbgimg}</td>
      <td class=tablerow2>����ͼƬ</td>
      <td class=tablerow2>{$fengmianlogo}</td>
      <td class=tablerow2>��վLOGO</td>
    </tr>
	<tr>
    <td class=tablerow2>{$fengmianbgmidi}</td>
      <td class=tablerow2>��������</td>
      <td class=tablerow2></td>
      <td class=tablerow2></td>
    </tr>

	<tr>
    <td class=tablerow2>     &lt;script src="{$InstallDir}inc/channel.js" type="text/javascript">&lt;/script></td>                    
      <td class=tablerow2>Ƶ���б�����</td>
      <td class=tablerow2>&lt;script src="{$InstallDir}count.asp" type="text/javascript">&lt;/script></td>                        
      <td class=tablerow2>������</td>
    </tr>





	</table>
	<br>
 <table border=0 align=center cellpadding=3 cellspacing=1 class=tableborder >
          <form name="myform" action="?action=addfengmian" method="post">
            <tbody>   
              
              <tr >
                <td width="190" class=tablerow2><strong>ģ������</strong></td>
                <td width="598" class=tablerow2>
                <input name="name" type="text" id="name" size="50" value=''><font color=red>*</font>

                </td>
              </tr>
              
              <tr class="tdbg" >
                <td width="190" class=tablerow2><strong>ģ���ţ�</strong></td>
                <td width="598" class=tablerow2>                                                              
                    <input name="bh" type="text" id="bh" size="50" value=''><font color=red>*</font>

                </td>                                                                                                              
              </tr>
              <tr class="tdbg" >
                <td width="190" class=tablerow2><strong>����Ŀ¼��</strong></td>
                <td width="598" class=tablerow2>                                                              
                    <input name="usedir" type="text" id="usedir" size="50" value=''>
                </td>                                                                                                              
              </tr>

              <tr class="tdbg">
                <td width="190" class=tablerow2><strong>����ͼ��</strong></td>
                <td width="598" class=tablerow2>                                                              
                    <input name="imageurl" type="text" id="ImageUrl" size="50" value=''>
                    <input type='button' name='selectpic' value='�����ϴ�ͼƬ��ѡ��' onclick='SelectPhoto()' class=button>
                    <iframe name="image" frameborder=0 width='100%' height=45 scrolling=no src=upload.asp?ChannelID=-1></iframe>
                </td>                                                                                                              
              </tr>
               
         		<tr class="tdbg">
                <td width="190" class=tablerow2><strong>����ͼƬ��</strong></td>
                <td width="598" class=tablerow2>                                                              
                    <input name="bgimg" type="text" id="bgimg" size="50" value=''>
                </td>                                                                                                              
              </tr>
                <tr class="tdbg">
                <td width="190" class=tablerow2><strong>�������֣�</strong></td>
                <td width="598" class=tablerow2>                                                              
                    <input name="bgmidi" type="text" id="bgmidi" size="50" value=''>
                </td>                                                                                                              
              </tr>

              <tr class="tdbg">
                <td width="190" class=tablerow2><strong>��վLOGO��</strong></td>
                <td width="598" class=tablerow2>                                                              
                    <input name="logo" type="text" id="logo" size="50" value=''>
                </td>                                                                                                              
              </tr>

              <tr class="tdbg">
                <td width="190" class=tablerow2><strong>ģ�����ݣ�(ģ�������ͼƬ��FLASHע��·�����ر�ע�⣬ͼƬ��FLASH�ļ��ϴ�����������������ѡ���ļ��ϴ�Ȼ���ٿ����ϴ����ļ���)</strong></td>
                <td width="598" class=tablerow2>  
                              <textarea name="content" id='content' style="display:none" rows="1" cols="20"></textarea>
					<iframe ID='HtmlEditor1' src='../editor/editor.asp?ChannelID=-1' frameborder=0 scrolling=no width='560' HEIGHT='350'></iframe>                                                                                                  
                </td>                                                                                                              
              </tr>
              
			<tr class="tdbg">
                <td width="190" class=tablerow2><strong>��ע��</strong></td>
                <td width="598" class=tablerow2>  
                <textarea rows="3" name="bz" cols="90"></textarea>                                                            
                </td>                                                                                                              
              </tr>

              
              
              <tr class="tdbg">
                <td align="middle" width="190" class=tablerow2>��</td>
                <td width="598" class=tablerow2><input type="submit" value="�ύ" name="Submit">��                                                                                                                                                                                                                                         
                  <input type="reset" value="����" name="Submit"></td>
              </tr>
              </FORM>
            </tbody>
          </table>


<%
end sub
Private Sub addfengmian()
		dim i,TextContent
		if Trim(Request.Form("name"))="" or Trim(Request.Form("bh"))="" or Trim(Request.Form("content"))="" then
			FoundErr = True
			ErrMsg = ErrMsg + "<li>����Ĳ�������!���������ݣ���Щ���ݲ���Ϊ�գ�</li>"

		else
		Set Rs = Server.CreateObject("ADODB.Recordset")
		if Request("id")<>"" then
			SQL = "select * from [ECCMS_fengmian] where id ="&LCase(Request("id"))&""
			Rs.Open SQL,Conn,1,3
		else
			SQL = "select * from [ECCMS_fengmian]"
			Rs.Open SQL,Conn,1,3
			Rs.Addnew
		end if
		
		rs("name")=Trim(Request.Form("name"))
		
		rs("bh")=Trim(Request.Form("bh"))
		
		if Trim(Request.Form("ImageUrl"))="" then
			rs("slt")=null
		else
			rs("slt")=Trim(Request.Form("ImageUrl"))
		end if	
		TextContent = ""
		For i = 1 To Request.Form("content").Count
			TextContent = TextContent & Request.Form("content")(i)
		Next

		rs("nr")=TextContent
		if Trim(Request.Form("bz"))="" then
		rs("bz")=null
		else

		rs("bz")=Trim(Request.Form("bz"))
		end if
		if Trim(Request.Form("bgimg"))="" then
		rs("bgimg")=null
		else

		rs("bgimg")=Trim(Request.Form("bgimg"))
		end if
		
		if Trim(Request.Form("logo"))="" then
		rs("logo")=null
		else

		rs("logo")=Trim(Request.Form("logo"))
		end if

		if Trim(Request.Form("bgmidi"))="" then
			rs("bgmidi")=null
		else

			rs("bgmidi")=Trim(Request.Form("bgmidi"))
		end if
		
		'����Ŀ¼
		if Trim(Request.Form("usedir"))="" then
			rs("usedir")=null
		else
		
			rs("usedir")=Trim(Request.Form("usedir"))

		end if
		
		
		Rs.update
		Rs.close:set Rs = Nothing
		Application.Contents.RemoveAll
		if Request("id")<>"" then

		Succeed("<li>��ϲ����ģ���޸ĳɹ���</li>")
		else
		Succeed("<li>��ϲ��������ģ��ɹ���</li>")
		end if
		end if

end sub
private sub editaddfengmianmoban()
Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL = "select * from [ECCMS_fengmian] where id ="&LCase(Request("id"))&""
	If IsSqlDataBase = 1 Then
		If CurrentPage > 100 Then
			Rs.Open SQL, Conn, 1, 1
		Else
			Set Rs = Conn.Execute(SQL)
		End If
	Else
		Rs.Open SQL, Conn, 1, 1
	End If
	If Rs.bof or Rs.EOF Then
		Response.Write "<table border=0 align=center cellpadding=3 cellspacing=1 class=tableborder><tr><td align=center colspan=5 class=TableRow2>��û��ѡ�����ģ�壡</td></tr></table>"
	else
		Response.Write "<table border=0 align=center cellpadding=3 cellspacing=1 class=tableborder><tr><td align=center colspan=5 class=TableRow2>��Ŀǰ�����޸ĵķ���ģ��Ϊ��(<font color=red>"& rs("name") &"</font>),ģ����Ϊ��<font color=red>"& rs("bh") &"</font></td></tr></table>"

	end if

%>
<script language= JavaScript>
function SelectPhoto(){
  var arr=showModalDialog('Admin_selFile.asp?ChannelID=-1&UploadDir=UploadPic', '', 'dialogWidth:800px; dialogHeight:600px; help: no; scroll: yes; status: no');
  if(arr!=null){
    var ss=arr.split('|');
    document.myform.ImageUrl.value=ss[0];
  }
}
</script>
<table border=0 align=center cellpadding=3 cellspacing=1 class=tableborder>
     <tr>
      <td>�����ǩ����</td>
     </tr>
     <tr>
      <td class=tablerow2>{$fengmianinstalldir}</td>
      <td class=tablerow2>����ģ������·��</td>
      <td class=tablerow2>{$Copyright}</td>
      <td class=tablerow2>��Ȩ��Ϣ</td>
    </tr>
	
	 <tr>
      <td class=tablerow2>{$InstallDir}</td>
      <td class=tablerow2>��վ����·��</td>
      <td class=tablerow2>{$fengmiancss}</td>
      <td class=tablerow2>��ʽ��Ϣ</td>
    </tr>

	<tr>
      <td class=tablerow2>{$fengmianbg}</td>
      <td class=tablerow2>ģ�屳����ɫ</td>
      <td class=tablerow2>{$WebSiteName}</td>
      <td class=tablerow2>��վ����</td>
    </tr>
	
	<tr>
      <td class=tablerow2>{$fengmianpic1}</td>
      <td class=tablerow2>��1��ͼƬ</td>
      <td class=tablerow2>{$fengmianflash1}</td>
      <td class=tablerow2>��1��FLASH</td>
    </tr>
	<tr>
    <td class=tablerow2>{$fengmianbgimg}</td>
      <td class=tablerow2>����ͼƬ</td>
      <td class=tablerow2>{$fengmianlogo}</td>
      <td class=tablerow2>��վLOGO</td>
    </tr>
	<tr>
    <td class=tablerow2>{$fengmianbgmidi}</td>
      <td class=tablerow2>��������</td>
      <td class=tablerow2></td>
      <td class=tablerow2></td>
    </tr>

	<tr>
    <td class=tablerow2>     &lt;script src="{$InstallDir}inc/channel.js" type="text/javascript">&lt;/script></td>                   
      <td class=tablerow2>Ƶ���б�����</td>
      <td class=tablerow2>&lt;script src="{$InstallDir}count.asp" type="text/javascript">&lt;/script></td>                       
      <td class=tablerow2>������</td>
    </tr>





	</table>
	<br>
 <table border=0 align=center cellpadding=3 cellspacing=1 class=tableborder >
          <form name="myform" action="?action=addfengmian&id=<%=rs("id")%>" method="post">
            <tbody>   
                       
              <tr >
                <td width="190" class=tablerow2><strong>ģ������</strong></td>
                <td width="598" class=tablerow2>
                <input name="name" type="text" id="name" size="50" value='<%=rs("name")%>'><font color=red>*</font>
                </td>
              </tr>
              
              <tr class="tdbg" >
                <td width="190" class=tablerow2><strong>ģ���ţ�</strong></td>
                <td width="598" class=tablerow2>                                                              
                    <input name="bh" type="text" id="bh" size="50" value='<%=rs("bh")%>'><font color=red>*</font>
                </td>                                                                                                              
              </tr>
                <tr class="tdbg" >
                <td width="190" class=tablerow2><strong>����Ŀ¼��</strong></td>
                <td width="598" class=tablerow2>                                                              
                    <input name="usedir" type="text" id="usedir" size="50" value='<%=rs("usedir")%>'>
                </td>                                                                                                              
              </tr>

              <tr class="tdbg">
                <td width="190" class=tablerow2><strong>����ͼ��</strong></td>
                <td width="598" class=tablerow2>                                                              
                    <input name="imageurl" type="text" id="ImageUrl" size="50" value='<%=rs("slt")%>'>
                    <input type='button' name='selectpic' value='�����ϴ�ͼƬ��ѡ��' onclick='SelectPhoto()' class=button>
                 <%
                    if rs("issystem")=1 then
                     	if Rs("slt")<>"" then 
                     		%>
                     		<img src=<%=enchiasp.InstallDir%><%=Rs("slt")%> width="100" height="75" align="center" border="1">  <br>
                     		<% 
                     	end if
                     else
                          if Rs("slt")<>"" then 
                          	%>
                          	<img src=<%=enchiasp.InstallDir%><%=Rs("slt")%> width="100" height="75" align="center" border="1"><br> 
                          	<% 
                          end if
                     end if
                    %>
                    <iframe name="image" frameborder=0 width='100%' height=45 scrolling=no src=upload.asp?ChannelID=-1></iframe>

                </td>                                                                                                              
              </tr>
               
         		<tr class="tdbg">
                <td width="190" class=tablerow2><strong>����ͼƬ��</strong></td>
                <td width="598" class=tablerow2>                                                              
                    <input name="bgimg" type="text" id="bgimg" size="50" value='<%=rs("bgimg")%>'>
                </td>                                                                                                              
              </tr>
                <tr class="tdbg">
                <td width="190" class=tablerow2><strong>�������֣�</strong></td>
                <td width="598" class=tablerow2>                                                              
                    <input name="bgmidi" type="text" id="bgmidi" size="50" value='<%=rs("bgmidi")%>'>
                </td>                                                                                                              
              </tr>

              <tr class="tdbg">
                <td width="190" class=tablerow2><strong>��վLOGO��</strong></td>
                <td width="598" class=tablerow2>                                                              
                    <input name="logo" type="text" id="logo" size="50" value='<%=rs("logo")%>'>
                </td>                                                                                                              
              </tr>

              <tr class="tdbg">
                <td width="190" class=tablerow2><strong>ģ�����ݣ�(ģ�������ͼƬ��FLASHע��·�����ر�ע�⣬ͼƬ��FLASH�ļ��ϴ�����������������ѡ���ļ��ϴ�Ȼ���ٿ����ϴ����ļ���)</strong></td>
                <td width="598" class=tablerow2>  
                              <textarea name="content" id='content' style="display:none" rows="1" cols="20"><%=rs("nr")%></textarea>  
					<iframe ID='HtmlEditor1' src='../editor/editor.asp?ChannelID=-1' frameborder=0 scrolling=no width='560' HEIGHT='350'></iframe>                                                                                                 
                </td>                                                                                                              
              </tr>
              
			<tr class="tdbg">
                <td width="190" class=tablerow2><strong>��ע��</strong></td>
                <td width="598" class=tablerow2>  
                <textarea rows="3" name="bz" cols="90"><%=rs("bz")%></textarea>                                                            
                </td>                                                                                                              
              </tr>

              
              
              <tr class="tdbg">
                <td align="middle" width="190" class=tablerow2>��</td>
                <td width="598" class=tablerow2><input type="submit" value="�ύ" name="Submit">��                                                                                                                                                                                                                                         
                  <input type="reset" value="����" name="Submit"></td>
              </tr>
              </FORM>
            </tbody>
          </table>


<%
end sub
private sub delfengmianmoban()
	If Request("id") = "" Then
		ErrMsg = "<li>��ѡ����ȷ��ϵͳ������</li>"
		Founderr = True
		Exit Sub
	End If
	enchiasp.Execute("Delete From [ECCMS_fengmian] Where id =" & Request("id"))
	Succeed("<li>��ϲ����ģ��ɾ���ɹ���</li>")

end sub
%>  








<%
Private Sub showpage()
	Dim n
	If totalnumber Mod maxperpage = 0 Then
		n = totalnumber \ maxperpage
	Else
		n = totalnumber \ maxperpage + 1
	End If
	Response.Write "<table cellspacing=1 width='100%' border=0><form method=Post action=?><tr><td align=center> " & vbCrLf
	If CurrentPage < 2 Then
		Response.Write "����ģ�� <font COLOR=#FF0000><strong>" & totalnumber & "</strong></font> ƪ&nbsp;�� ҳ&nbsp;��һҳ&nbsp;|&nbsp;"
	Else
		Response.Write "����ģ�� <font COLOR=#FF0000><strong>" & totalnumber & "</strong></font> ƪ&nbsp;<a href=?page=1>�� ҳ</a>&nbsp;"
		Response.Write "<a href=?page=" & CurrentPage - 1 &  ">��һҳ</a>&nbsp;|&nbsp;"
	End If
	If n - CurrentPage < 1 Then
		Response.Write "��һҳ&nbsp;β ҳ" & vbCrLf
	Else
		Response.Write "<a href=?page=" & (CurrentPage + 1) & ">��һҳ</a>"
		Response.Write "&nbsp;<a href=?page=" & n & ">β ҳ</a>" & vbCrLf
	End If
	Response.Write "&nbsp;ҳ�Σ�<strong><font color=red>" & CurrentPage & "</font>/" & n & "</strong>ҳ "
	Response.Write "&nbsp;ת����"
	Response.Write "<input name=page size=3 value='" & CurrentPage & "'> <input class=Button type=submit name=Submit value='ת��'>"
	Response.Write "</td></tr></FORM></table>" & vbCrLf
End Sub

%>




















































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
Dim bkfolder, bkdbname, fso, fso1
Dim Action
Action = LCase(Request("action"))

Select Case Action
	Case "backupdata" '��������
		If Not ChkAdmin("BackupData") Then
			Server.Transfer("showerr.asp")
			Request.End
		End If
		If request("act") = "Backup" Then
			If IsSqlDataBase = 1 Then
				Call BackupSqlDatabase()
			Else
				Call BackupDatabase()
			End If
		Else
			Call BackupData()
		End If
	Case "compressdata" 'ѹ������
		If Not ChkAdmin("CompressData") Then
			Server.Transfer("showerr.asp")
			Request.End
		End If
		If request("act") = "Compress" Then
			Call CompressDatabase()
		Else
			Call CompressData()
		End If
	Case "restoredata" '�ָ�����
		If Not ChkAdmin("RestoreData") Then
			Server.Transfer("showerr.asp")
			Request.End
		End If
		If request("act") = "Restore" Then
			If IsSqlDataBase = 1 Then
				Call RestoreSqlDatabase()
			Else
				Call RestoreDatabase
			End If
			Application.Contents.RemoveAll
		Else
			Call RestoreData()
		End If

	Case "spacesize" 'ϵͳ�ռ�ռ��
		If Not ChkAdmin("SpaceSize") Then
			Server.Transfer("showerr.asp")
			Request.End
		End If
		Call SpaceSize()

	Case Else
		Errmsg = ErrMsg + "<BR><li>ѡȡ��Ӧ�Ĳ�����"
		ReturnError(ErrMsg)

End Select
If FoundErr = True Then
	ReturnError(ErrMsg)
End If
Admin_footer
SaveLogInfo(AdminName)
CloseConn
'====================ϵͳ�ռ�ռ��=======================
Sub SpaceSize()
	On Error Resume Next
%>
		<table border="0"  cellspacing="1" cellpadding="5" height="1" align=center width="95%" class="tableBorder">			<tr>
  					<th height=25>
  					&nbsp;&nbsp;ϵͳ�ռ�ռ�����
  					</th>
  				</tr> 	
 				<tr>
 					<td class="TableRow1"> 			
 			<blockquote>
 			<br> 			
 			��������ռ�ÿռ䣺&nbsp;<img src="images/bar1.gif" width=<%=drawbar("../Database")%> height=10>&nbsp;<%showSpaceinfo("../database")%><br><br>
 			��������ռ�ÿռ䣺&nbsp;<img src="images/bar1.gif" width=<%=drawbar("databackup")%> height=10>&nbsp;<%showSpaceinfo("databackup")%><br><br>
			��̨����ռ�ÿռ䣺&nbsp;<img src="images/bar1.gif" width=<%=drawbar("./")%> height=10>&nbsp;<%showSpaceinfo("./")%><br><br>
 			���Ƶ��ռ�ÿռ䣺&nbsp;<img src="images/bar1.gif" width=<%=drawbar("../Soft")%> height=10>&nbsp;<%showSpaceinfo("../Soft")%><br><br>
			����Ƶ��ռ�ÿռ䣺&nbsp;<img src="images/bar1.gif" width=<%=drawbar("../Article")%> height=10>&nbsp;<%showSpaceinfo("../Article")%><br><br>
 			ģ��ͼƬռ�ÿռ䣺&nbsp;<img src="images/bar1.gif" width=<%=drawbar("../skin")%> height=10>&nbsp;<%showSpaceinfo("../skin")%><br><br>
 			ͼƬ�ļ�ռ�ÿռ䣺&nbsp;<img src="images/bar1.gif" width=<%=drawbar("../images")%> height=10>&nbsp;<%showSpaceinfo("../images")%><br><br>
 			�û��ļ�ռ�ÿռ䣺&nbsp;<img src="images/bar1.gif" width=<%=drawbar("../user")%> height=10>&nbsp;<%showSpaceinfo("../user")%><br><br>
			�ϴ��ļ�ռ�ÿռ䣺&nbsp;<img src="images/bar1.gif" width=<%=drawbar("../UploadFile")%> height=10>&nbsp;<%showSpaceinfo("../UploadFile")%><br><br>
 			ϵͳռ�ÿռ��ܼƣ�&nbsp;<img src="images/bar2.gif" width=400 height=10> <%showspecialspaceinfo("All")%>
 			</blockquote> 	
 					</td>
 				</tr>
 			</table>
<%
End Sub

Sub SQLUserReadme()
%>
		<table border="0"  cellspacing="1" cellpadding="5" height="1" align=center width="95%" class="tableBorder">			<tr>
  					<th height=25>
  					&nbsp;&nbsp;SQL���ݿ����ݴ���˵��
  					</th>
  				</tr> 	
 				<tr>
 					<td class="TableRow1"> 			
 			<blockquote>
<B>һ���������ݿ�</B>
<BR><BR>
1����SQL��ҵ���������ڿ���̨��Ŀ¼�����ε㿪Microsoft SQL Server<BR>
2��SQL Server��-->˫������ķ�����-->˫�������ݿ�Ŀ¼<BR>
3��ѡ��������ݿ����ƣ���ϵͳ���ݿ�enchiasp��-->Ȼ�������˵��еĹ���-->ѡ�񱸷����ݿ�<BR>
4������ѡ��ѡ����ȫ���ݣ�Ŀ���еı��ݵ����ԭ����·����������ѡ�����Ƶ�ɾ����Ȼ�����ӣ����ԭ��û��·����������ֱ��ѡ����ӣ�����ָ��·�����ļ�����ָ�����ȷ�����ر��ݴ��ڣ����ŵ�ȷ�����б���
<BR><BR>
<B>������ԭ���ݿ�</B><BR><BR>
1����SQL��ҵ���������ڿ���̨��Ŀ¼�����ε㿪Microsoft SQL Server<BR>
2��SQL Server��-->˫������ķ�����-->��ͼ�������½����ݿ�ͼ�꣬�½����ݿ����������ȡ<BR>
3������½��õ����ݿ����ƣ���ϵͳ���ݿ�enchiasp��-->Ȼ�������˵��еĹ���-->ѡ��ָ����ݿ�<BR>
4���ڵ������Ĵ����еĻ�ԭѡ����ѡ����豸-->��ѡ���豸-->�����-->Ȼ��ѡ����ı����ļ���-->��Ӻ��ȷ�����أ���ʱ���豸��Ӧ�ó������ղ�ѡ������ݿⱸ���ļ��������ݺ�Ĭ��Ϊ1���������ͬһ���ļ�������α��ݣ����Ե�����ݺ��ԱߵĲ鿴���ݣ��ڸ�ѡ����ѡ�����µ�һ�α��ݺ��ȷ����-->Ȼ�����Ϸ������Աߵ�ѡ�ť<BR>
5���ڳ��ֵĴ�����ѡ�����������ݿ���ǿ�ƻ�ԭ���Լ��ڻָ����״̬��ѡ��ʹ���ݿ���Լ������е��޷���ԭ����������־��ѡ��ڴ��ڵ��м䲿λ�Ľ����ݿ��ļ���ԭΪ����Ҫ������SQL�İ�װ�������ã�Ҳ����ָ���Լ���Ŀ¼�����߼��ļ�������Ҫ�Ķ������������ļ���Ҫ���������ָ��Ļ���������Ķ���������SQL���ݿ�װ��D:\Program Files\Microsoft SQL Server\MSSQL\Data����ô�Ͱ������ָ�������Ŀ¼������ظĶ��Ķ������������ļ�����øĳ�����ǰ�����ݿ�������ԭ����enchicms_data.mdf�����ڵ����ݿ���enchiasp���͸ĳ�enchiasp_data.mdf������־�������ļ���Ҫ���������ķ�ʽ����صĸĶ�����־���ļ�����*_log.ldf��β�ģ�������Ļָ�Ŀ¼�������������ã�ǰ���Ǹ�Ŀ¼������ڣ���������ָ��d:\sqldata\enchicms_data.mdf����d:\sqldata\enchicms_log.ldf��������ָ�������<BR>
6���޸���ɺ󣬵�������ȷ�����лָ�����ʱ�����һ������������ʾ�ָ��Ľ��ȣ��ָ���ɺ�ϵͳ���Զ���ʾ�ɹ������м���ʾ�������¼����صĴ������ݲ�ѯ�ʶ�SQL�����Ƚ���Ϥ����Ա��һ��Ĵ����޷���Ŀ¼��������ļ����ظ������ļ���������߿ռ䲻���������ݿ�����ʹ���еĴ������ݿ�����ʹ�õĴ��������Գ��Թر����й���SQL����Ȼ�����´򿪽��лָ��������������ʾ����ʹ�õĴ�����Խ�SQL����ֹͣȻ�����𿴿����������������Ĵ���һ�㶼�ܰ��մ�����������Ӧ�Ķ��󼴿ɻָ�<BR><BR>

<B>�����������ݿ�</B><BR><BR>
һ������£�SQL���ݿ�����������ܴܺ�̶��ϼ�С���ݿ��С������Ҫ������������־��С��Ӧ�����ڽ��д˲����������ݿ���־����<BR>
1���������ݿ�ģʽΪ��ģʽ����SQL��ҵ���������ڿ���̨��Ŀ¼�����ε㿪Microsoft SQL Server-->SQL Server��-->˫������ķ�����-->˫�������ݿ�Ŀ¼-->ѡ��������ݿ����ƣ���ϵͳ���ݿ�enchiasp��-->Ȼ�����Ҽ�ѡ������-->ѡ��ѡ��-->�ڹ��ϻ�ԭ��ģʽ��ѡ�񡰼򵥡���Ȼ��ȷ������<BR>
2���ڵ�ǰ���ݿ��ϵ��Ҽ��������������е��������ݿ⣬һ�������Ĭ�����ò��õ�����ֱ�ӵ�ȷ��<BR>
3��<font color=blue>�������ݿ���ɺ󣬽��齫�������ݿ�������������Ϊ��׼ģʽ����������ͬ��һ�㣬��Ϊ��־��һЩ�쳣����������ǻָ����ݿ����Ҫ����</font>
<BR><BR>

<B>�ġ��趨ÿ���Զ��������ݿ�</B><BR><BR>
<font color=red>ǿ�ҽ������������û����д˲�����</font><BR>
1������ҵ���������ڿ���̨��Ŀ¼�����ε㿪Microsoft SQL Server-->SQL Server��-->˫������ķ�����<BR>
2��Ȼ�������˵��еĹ���-->ѡ�����ݿ�ά���ƻ���<BR>
3����һ��ѡ��Ҫ�����Զ����ݵ�����-->��һ�����������Ż���Ϣ������һ�㲻����ѡ��-->��һ��������������ԣ�Ҳһ�㲻ѡ��<BR>
4����һ��ָ�����ݿ�ά���ƻ���Ĭ�ϵ���1�ܱ���һ�Σ��������ѡ��ÿ�챸�ݺ��ȷ��<BR>
5����һ��ָ�����ݵĴ���Ŀ¼��ѡ��ָ��Ŀ¼������������D���½�һ��Ŀ¼�磺d:\databak��Ȼ��������ѡ��ʹ�ô�Ŀ¼������������ݿ�Ƚ϶����ѡ��Ϊÿ�����ݿ⽨����Ŀ¼��Ȼ��ѡ��ɾ�����ڶ�����ǰ�ı��ݣ�һ���趨4��7�죬�⿴���ľ��屸��Ҫ�󣬱����ļ���չ��һ�㶼��bak����Ĭ�ϵ�<BR>
6����һ��ָ��������־���ݼƻ�����������Ҫ��ѡ��-->��һ��Ҫ���ɵı���һ�㲻��ѡ��-->��һ��ά���ƻ���ʷ��¼�������Ĭ�ϵ�ѡ��-->��һ�����<BR>
7����ɺ�ϵͳ�ܿ��ܻ���ʾSql Server Agent����δ�������ȵ�ȷ����ɼƻ��趨��Ȼ���ҵ��������ұ�״̬���е�SQL��ɫͼ�꣬˫���㿪���ڷ�����ѡ��Sql Server Agent��Ȼ�������м�ͷ��ѡ���·��ĵ�����OSʱ�Զ���������<BR>
8�����ʱ�����ݿ�ƻ��Ѿ��ɹ��������ˣ�������������������ý����Զ�����
<BR><BR>
�޸ļƻ���<BR>
1������ҵ���������ڿ���̨��Ŀ¼�����ε㿪Microsoft SQL Server-->SQL Server��-->˫������ķ�����-->����-->���ݿ�ά���ƻ�-->�򿪺�ɿ������趨�ļƻ������Խ����޸Ļ���ɾ������
<BR><BR>
<B>�塢���ݵ�ת�ƣ��½����ݿ��ת�Ʒ�������</B><BR><BR>
һ������£����ʹ�ñ��ݺͻ�ԭ����������ת�����ݣ�����������£������õ��뵼���ķ�ʽ����ת�ƣ�������ܵľ��ǵ��뵼����ʽ�����뵼����ʽת������һ�����þ��ǿ������������ݿ���Ч�������������С�����������ݿ�Ĵ�С��������Ĭ��Ϊ����SQL�Ĳ�����һ�����˽⣬��������еĲ��ֲ�������⣬������ѯ���������Ա���߲�ѯ��������<BR>
1����ԭ���ݿ�����б��洢���̵�����һ��SQL�ļ���������ʱ��ע����ѡ����ѡ���д�����ű��ͱ�д�����������Ĭ��ֵ�ͼ��Լ���ű�ѡ��<BR>
2���½����ݿ⣬���½����ݿ�ִ�е�һ������������SQL�ļ�<BR>
3����SQL�ĵ��뵼����ʽ���������ݿ⵼��ԭ���ݿ��е����б�����<BR>
 			</blockquote> 	
 					</td>
 				</tr>
 			</table>
<%
End Sub

'====================�ָ����ݿ�=========================

Sub RestoreData()
	If IsSqlDataBase = 1 Then
%>
	<table border="0"  cellspacing="1" cellpadding="5" height="1" align=center width="95%" class="tableBorder">
  				<tr>
  					<th height=25 >
  					&nbsp;&nbsp;�ָ�SQL����
  					</th>
  				</tr>
  				<form method="post" action="?action=RestoreData&act=Restore">
  				<tr>
  					<td height=100 class="TableRow1">
  						&nbsp;&nbsp;
						�ָ�SQL���ݿ����ƣ�<input type=text size=25 name=SqlDataName value="<%=SqlDatabaseName%>"><BR>&nbsp;&nbsp;
						SQL���ݿ��û����ƣ�<input type=text size=25 name=SqlUserID value="<%=SqlUsername%>">&nbsp;����������SQL���ݿ��û���<BR>&nbsp;&nbsp;
						SQL���ݿ��û����룺<input type=Password size=25 name=SqlUserPass value="<%=SqlPassword%>">&nbsp;����������SQL���ݿ���������<BR>&nbsp;&nbsp;
						SQL���ݿ����������<input type=text size=25 name=SqlServer value="<%=SqlLocalName%>">&nbsp;���ӷ���������������local�������IP��<BR>&nbsp;&nbsp;
						����SQL���ݿ�Ŀ¼��<input type=text size=25 name=BackupSqlDir value="Databackup">&nbsp;�����������ݵ����ݿ�Ŀ¼<BR>&nbsp;&nbsp;
						����SQL���ݿ����ƣ�<input type=text size=25 name=BackupSqlName  value="$1.bak">&nbsp;�����������ݵ����ݿ���<BR>
						&nbsp;&nbsp;<input type=submit value="�ָ����ݿ�" class=Button><br>
  						-----------------------------------------------------------------------------------------<br>
  						&nbsp;&nbsp;������������������ָ����ķ������ݣ���ȷ��������ݿ��û������Ȩ�޲��ָܻ���<br>
  						&nbsp;&nbsp;ע�⣺���ݿ�ʹ���п����޷��ָ�	</font>
  					</td>
  				</tr>	
  				</form>
  	</table>
<%
	Else
%>
<table border="0"  cellspacing="1" cellpadding="5" height="1" align=center width="95%" class="tableBorder"		<tr>
	<th height=25 >
   					&nbsp;&nbsp;�ָ�ϵͳ���� ( ��ҪFSO֧�֣�FSO��ذ����뿴΢����վ )
  					</th>
  				</tr>
				<form method="post" action="?action=RestoreData&act=Restore">
  				
  				<tr>
  					<td height=100 class="TableRow1">
  						&nbsp;&nbsp;�������ݿ�·��(���)��<input type=text size=45 name=DBpath value="DataBackup\enchicms_Backup.MDB">&nbsp;&nbsp;<BR>
  						&nbsp;&nbsp;Ŀ�����ݿ�·��(���)��<input type=text size=45 name=backpath value="<%=db%>"><BR>&nbsp;&nbsp;��д����ǰʹ�õ����ݿ�·�����粻�븲�ǵ�ǰ�ļ���������������ע��·���Ƿ���ȷ����Ȼ���޸�conn.asp�ļ������Ŀ���ļ����͵�ǰʹ�����ݿ���һ�µĻ��������޸�conn.asp�ļ�<BR>
						&nbsp;&nbsp;<input type=submit value="�ָ����ݿ�" onclick="{if(confirm('��ȷ��Ҫ�ָ����ݿ���?')){return true;}return false;}" class=Button> <br>
  						-----------------------------------------------------------------------------------------<br>
  						&nbsp;&nbsp;��������д����������ݿ�·��ȫ�����������Ĭ�ϱ������ݿ��ļ�ΪDataBackup\enchicms_Backup.MDB���밴�����ı����ļ������޸ġ�<br>
  						&nbsp;&nbsp;������������������������ķ������ݣ��Ա�֤�������ݰ�ȫ��<br>
  						&nbsp;&nbsp;ע�⣺����·��������������ռ��Ŀ¼�����·��</font>
  					</td>
  				</tr>	
  				</form>
  			</table>
<%
	End If
End Sub
'====================�������ݿ�=========================
Sub BackupData()
	If IsSqlDataBase = 1 Then
%>
	<table border="0"  cellspacing="1" cellpadding="5" height="1" align=center width="95%" class="tableBorder">
  				<tr>
  					<th height=25 >
  					&nbsp;&nbsp;����SQL����
  					</th>
  				</tr>
  				<form method="post" action="?action=BackupData&act=Backup">
  				<tr>
  					<td height=100 class="TableRow1">
  						&nbsp;&nbsp;
						��ǰSQL���ݿ����ƣ�<input type=text size=25 name=SqlDataName value="<%=SqlDatabaseName%>"><BR>&nbsp;&nbsp;
						SQL���ݿ��û����ƣ�<input type=text size=25 name=SqlUserID value="<%=SqlUsername%>">&nbsp;����������SQL���ݿ��û���<BR>&nbsp;&nbsp;
						SQL���ݿ��û����룺<input type=Password size=25 name=SqlUserPass value="<%=SqlPassword%>">&nbsp;����������SQL���ݿ���������<BR>&nbsp;&nbsp;
						SQL���ݿ����������<input type=text size=25 name=SqlServer value="<%=SqlLocalName%>">&nbsp;���ӷ���������������local�������IP��<BR>&nbsp;&nbsp;
						����SQL���ݿ�Ŀ¼��<input type=text size=25 name=BackupSqlDir value="Databackup">&nbsp;��Ŀ¼�����ڣ������Զ�����<BR>&nbsp;&nbsp;
						����SQL���ݿ����ƣ�<input type=text size=25 name=BackupSqlName  value="$1.bak">&nbsp;��ʹ��Ĭ�ϵ�($1)������,ϵͳ���Զ�������ʱ�䴴����������<BR>
						&nbsp;&nbsp;<input type=submit value="�������ݿ�" class=Button><br>
  						-----------------------------------------------------------------------------------------<br>
  						&nbsp;&nbsp;������������������������ķ������ݣ��Ա�֤�������ݰ�ȫ��<br>
  						&nbsp;&nbsp;ע�⣺����·��������������ռ����Ŀ¼�����·��				</font>
  					</td>
  				</tr>	
  				</form>
  	</table>
<%
Else
%>
	<table border="0"  cellspacing="1" cellpadding="5" height="1" align=center width="95%" class="tableBorder">
  				<tr>
  					<th height=25 >
  					&nbsp;&nbsp;����ϵͳ���� ( ��ҪFSO֧�֣�FSO��ذ����뿴΢����վ )
  					</th>
  				</tr>
  				<form method="post" action="?action=BackupData&act=Backup">
  				<tr>
  					<td height=100 class="TableRow1">
  						&nbsp;&nbsp;
						��ǰ���ݿ�·��(���·��)��<input type=text size=45 name=DBpath size=45 value="<%=db%>"><BR>&nbsp;&nbsp;
						�������ݿ�Ŀ¼(���·��)��<input type=text size=20 name=bkfolder size=45 value=Databackup>&nbsp;��Ŀ¼�����ڣ������Զ�����<BR>&nbsp;&nbsp;
						�������ݿ�����(��д����)��<input type=text size=20 name=bkDBname size=45 value=enchicms_Backup.MDB>&nbsp;�籸��Ŀ¼�и��ļ��������ǣ���û�У����Զ�����<BR>
						&nbsp;&nbsp;<input type=submit value="�������ݿ�" class=Button><br>
  						-----------------------------------------------------------------------------------------<br>
  						&nbsp;&nbsp;��������д����������ݿ�·��ȫ�����������Ĭ�����ݿ��ļ�Ϊ<%=db%>��<B>��һ��������Ĭ�����������������ݿ�</B><br>
  						&nbsp;&nbsp;������������������������ķ������ݣ��Ա�֤�������ݰ�ȫ��<br>
  						&nbsp;&nbsp;ע�⣺����·��������������ռ����Ŀ¼�����·��				</font>
  					</td>
  				</tr>	
  				</form>
  	</table>
<%
End If
End Sub

Sub BackupDatabase()
	Dbpath = request.Form("Dbpath")
	If InStr(Dbpath, ":") = 0 Then
		Dbpath = Server.MapPath(Dbpath)
	Else
		Dbpath = Dbpath
	End If
	bkfolder = request.Form("bkfolder")
	bkdbname = request.Form("bkdbname")
	Set Fso = server.CreateObject("scripting.filesystemobject")
	If fso.FileExists(dbpath) Then
		If CheckDir(bkfolder) = True Then
			fso.CopyFile dbpath, bkfolder& "\"& bkdbname
		Else
			MakeNewsDir bkfolder
			fso.CopyFile dbpath, bkfolder& "\"& bkdbname
		End If
		Succeed("�������ݿ�ɹ��������ݵ����ݿ�·��Ϊ" &bkfolder& "\"& bkdbname)
	Else
		FoundErr = True
		ErrMsg = "�Ҳ���������Ҫ���ݵ��ļ���"
		Exit Sub
	End If
End Sub
Sub RestoreDatabase()
	Dim backpath,Dbpath
	Dbpath = request.Form("Dbpath")
	backpath = request.Form("backpath")
	If dbpath = "" Then
		FoundErr = True
		ErrMsg = "��������Ҫ�ָ��ɵ����ݿ�ȫ��"
		Exit Sub
	End If
	If InStr(Dbpath, ":") = 0 Then
		Dbpath = Server.MapPath(Dbpath)
	Else
		Dbpath = Dbpath
	End If
	If InStr(backpath, ":") = 0 Then
		backpath = Server.MapPath(backpath)
	Else
		backpath = backpath
	End If
	Set Fso = server.CreateObject("scripting.filesystemobject")
	If fso.FileExists(dbpath) Then
		fso.CopyFile Dbpath, Backpath
		Succeed("�ɹ��ָ����ݣ�")
	Else
		FoundErr = True
		ErrMsg = "����Ŀ¼�²������ı����ļ���"
		Exit Sub
	End If
End Sub
'------------------���ĳһĿ¼�Ƿ����-------------------
Function CheckDir(FolderPath)
	folderpath = Server.MapPath(".")&"\"&folderpath
	Set fso1 = CreateObject("Scripting.FileSystemObject")
	If fso1.FolderExists(FolderPath) Then
		'����
		CheckDir = True
	Else
		'������
		CheckDir = False
	End If
	Set fso1 = Nothing
End Function
'-------------����ָ����������Ŀ¼-----------------------
Function MakeNewsDir(foldername)
	Dim f
	Set fso1 = CreateObject("Scripting.FileSystemObject")
	Set f = fso1.CreateFolder(foldername)
	MakeNewsDir = True
	Set fso1 = Nothing
End Function
'====================ѹ�����ݿ� =========================
Sub CompressData()

	If IsSqlDataBase =1 Then
		SQLUserReadme()
		Exit Sub
	End If
%>
<table border="0"  cellspacing="1" cellpadding="5" height="1" align=center width="95%" class="tableBorder">
<tr>
<th height=25 >
&nbsp;&nbsp;ѹ�����ݿ� ( ��ҪFSO֧�֣�FSO��ذ����뿴΢����վ )
</th>
<form action="?action=CompressData&act=Compress" method="post">
<tr>
<td class="TableRow1" height=25><b>ע�⣺</b><br>�������ݿ��������·��,�����������ݿ����ƣ�����ʹ�������ݿⲻ��ѹ������ѡ�񱸷����ݿ����ѹ�������� </td>
</tr>
<tr>
<td class="TableRow1">ѹ�����ݿ⣺<input type="text" name="dbpath" size=45 value=<%=db%>>&nbsp;
<input type="submit" value="��ʼѹ��" class=Button></td>
</tr>
<tr>
<td class="TableRow1"><input type="checkbox" name="boolIs97" value="True">���ʹ�� Access 97 ���ݿ���ѡ��
(Ĭ��Ϊ Access 2000 ���ݿ�)<br><br></td>
</tr>
<form>
</table>
<%
End Sub

Sub CompressDatabase()
	Dim dbpath, boolIs97
	dbpath = request("dbpath")
	boolIs97 = request("boolIs97")

	If dbpath <> "" Then
		If InStr(Dbpath, ":") = 0 Then
			Dbpath = Server.MapPath(Dbpath)
		Else
			Dbpath = Dbpath
		End If
		Response.Write(CompactDB(dbpath, boolIs97))
	Else
		FoundErr = True
		ErrMsg = "������Ҫѹ�������ݿ�·����"
		Exit Sub
	End If
End Sub
'=====================ѹ������=========================
Function CompactDB(dbPath, boolIs97)
	Dim fso, Engine, strDBPath, JET_3X
	strDBPath = Left(dbPath, instrrev(DBPath, "\"))
	Set fso = CreateObject("Scripting.FileSystemObject")

	If fso.FileExists(dbPath) Then
		fso.CopyFile dbpath, strDBPath & "temp.mdb"
		Set Engine = CreateObject("JRO.JetEngine")

		If boolIs97 = "True" Then
			Engine.CompactDatabase "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBPath & "temp.mdb", _
				"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBPath & "temp1.mdb;" _
				& "Jet OLEDB:Engine Type=" & JET_3X
		Else
			Engine.CompactDatabase "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBPath & "temp.mdb", _
				"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBPath & "temp1.mdb"
		End If

		fso.CopyFile strDBPath & "temp1.mdb", dbpath
		fso.DeleteFile(strDBPath & "temp.mdb")
		fso.DeleteFile(strDBPath & "temp1.mdb")
		Set fso = Nothing
		Set Engine = Nothing
		Succeed("������ݿ�, " & dbpath & ", �Ѿ�ѹ���ɹ�!")
	Else
		ReturnError("���ݿ����ƻ�·������ȷ. ������!")
	End If

End Function
'=====================ϵͳ�ռ����=========================
Sub ShowSpaceInfo(drvpath)
	Dim fso, d, Size, showsize
	Set fso = server.CreateObject("scripting.filesystemobject")
	drvpath = server.mappath(drvpath)
	Set d = fso.GetFolder(drvpath)
	Size = d.Size
	showsize = Size & "&nbsp;Byte"
	If Size>1024 Then
		Size = (Size / 1024)
		showsize = FormatNumber(Size, 2) & "&nbsp;KB"
	End If
	If Size>1024 Then
		Size = (Size / 1024)
		showsize = FormatNumber(Size, 2) & "&nbsp;MB"
	End If
	If Size>1024 Then
		Size = (Size / 1024)
		showsize = FormatNumber(Size, 2) & "&nbsp;GB"
	End If
	response.Write "<font face=verdana>" & showsize & "</font>"
End Sub

Sub Showspecialspaceinfo(method)
	dim fso,d,fc,f1,size,showsize,drvpath 		
	set fso=server.CreateObject("Scr"&"ipting"&".Fil"&"eSy"&"stemOb"&"ject")
	drvpath=server.mappath("../inc")
	drvpath=left(drvpath,(instrrev(drvpath,"\")-1))
	set d=fso.getfolder(drvpath) 	
	if method="All" then 		
		size=d.size
	elseif method="Program" then
		set fc=d.Files
		for each f1 in fc
			size=size+f1.size
		next	
	end if			
	showsize=size & "&nbsp;Byte" 
	if size>1024 then
	   size=(size\1024)
	   showsize=size & "&nbsp;KB"
	end if
	if size>1024 then
	   size=(size/1024)
	   showsize=formatnumber(size,2) & "&nbsp;MB"		
	end if
	if size>1024 then
	   size=(size/1024)
	   showsize=formatnumber(size,2) & "&nbsp;GB"	   
	end if   
	response.write "<font face=verdana>" & showsize & "</font>"

	

	'drvpath = server.mappath("../../../")

End Sub

Function Drawbar(drvpath)
	Dim fso, drvpathroot, d, Size, TotalSize, barsize
	Set fso = server.CreateObject("scripting.filesystemobject")
	drvpathroot = server.mappath("../pic")
	drvpathroot = Left(drvpathroot, (instrrev(drvpathroot, "\") -1))
	Set d = fso.GetFolder(drvpathroot)
	TotalSize = d.Size

	drvpath = server.mappath(drvpath)
	Set d = fso.GetFolder(drvpath)
	Size = d.Size

	barsize = CDbl((Size / TotalSize) * 400)
	Drawbar = barsize
End Function

Function Drawspecialbar()
	Dim fso, drvpathroot, d, fc, f1, Size, TotalSize, barsize
	Set fso = server.CreateObject("scripting.filesystemobject")
	drvpathroot = server.mappath("../pic")
	drvpathroot = Left(drvpathroot, (instrrev(drvpathroot, "\") -1))
	Set d = fso.GetFolder(drvpathroot)
	TotalSize = d.Size

	Set fc = d.Files
	For Each f1 in fc
		Size = Size + f1.Size
	Next

	barsize = CDbl((Size / TotalSize) * 400)
	Drawspecialbar = barsize
End Function

Sub CheckSql()
	If Trim(Request.Form("SqlDataName")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg & "<li>������SQL���ݿ�����</li>"
	End If
	If Trim(Request.Form("SqlUserPass")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg & "<li>������SQL���ݿ��û����룡</li>"
	End If
	If Trim(Request.Form("SqlUserID")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg & "<li>������SQL���ݿ��û����ƣ�</li>"
	End If
	If Trim(Request.Form("SqlServer")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg & "<li>������SQL���ݿ���������������local�������IP����</li>"
	End If
	If Trim(Request.Form("BackupSqlName")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg & "<li>������SQL���ݿⱸ�����ƣ�</li>"
	End If
	If Trim(Request.Form("BackupSqlDir")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg & "<li>������SQL���ݿⱸ��Ŀ¼��</li>"
	End If
End Sub
'====================����SQL���ݿ�=========================
Sub BackupSqlDatabase()
	On Error Resume Next
	Dim SqlDataName, SqlUserPass, SqlUserID, SqlServer, SqlLoginTimeout
	Dim srv, bak, BackupFilePath, BackupSqlDir, BackupSqlName,BackupFileName
	SqlDataName = Trim(Request.Form("SqlDataName"))
	SqlUserPass = Trim(Request.Form("SqlUserPass"))
	SqlUserID = Trim(Request.Form("SqlUserID"))
	SqlServer = Trim(Request.Form("SqlServer"))
	BackupSqlDir = Trim(Request.Form("BackupSqlDir"))
	BackupSqlName = Trim(Request.Form("BackupSqlName"))
	SqlLoginTimeout = 20 '��½��ʱ
	CheckSql
	If FoundErr = True Then Exit Sub
	If CheckDir(BackupSqlDir) = False Then
		MakeNewsDir BackupSqlDir
	End If
	BackupFileName = SqlDataName & "_" & Replace(FormatDateTime(now,2), "-", "") & "_" & Replace(FormatDateTime(now,3), ":", "")
	BackupFilePath = BackupSqlDir & "\" & BackupSqlName
	BackupFilePath = Replace(BackupFilePath, "$1", BackupFileName)
	Set srv = Server.CreateObject("SQLDMO.SQLServer")
	srv.LoginTimeout = SqlLoginTimeout
	srv.Connect SqlServer, SqlUserID, SqlUserPass
	Set bak = Server.CreateObject("SQLDMO.Backup")
	bak.Database = SqlDataName
	'bak.Devices = Files
	bak.Files = BackupFilePath
	bak.SQLBackup srv
	If Err.Number>0 Then
		Response.Write Err.Number & "<font color=red><br>"
		Response.Write Err.Description & "</font>"
	End If
	Set srv = Nothing
	Set bak = Nothing
	Succeed("<li>SQL���ݿⱸ�ݳɹ���</li>")
End Sub
'====================�ָ�SQL���ݿ�=========================
Sub RestoreSqlDatabase()
	On Error Resume Next
	Dim SqlDataName, SqlUserPass, SqlUserID, SqlServer, SqlLoginTimeout
	Dim srv, rest, BackupFilePath, BackupSqlDir, BackupSqlName, FSO
	SqlDataName = Trim(Request.Form("SqlDataName"))
	SqlUserPass = Trim(Request.Form("SqlUserPass"))
	SqlUserID = Trim(Request.Form("SqlUserID"))
	SqlServer = Trim(Request.Form("SqlServer"))
	BackupSqlDir = Trim(Request.Form("BackupSqlDir"))
	BackupSqlName = Trim(Request.Form("BackupSqlName"))
	SqlLoginTimeout = 20 '��½��ʱ
	CheckSql
	If FoundErr = True Then Exit Sub
	BackupFilePath = BackupSqlDir & "/" & BackupSqlName
	BackupFilePath = Replace(BackupFilePath, "$1", SqlDataName)
	BackupFilePath = Server.MapPath(BackupFilePath)
	Set FSO = Server.CreateObject("scripting.filesystemobject")
	If FSO.FileExists(BackupFilePath) Then
		Set srv = Server.CreateObject("SQLDMO.SQLServer")
		srv.LoginTimeout = SqlLoginTimeout
		srv.Connect SqlServer, SqlUserID, SqlUserPass
		Set rest = Server.CreateObject("SQLDMO.Restore")
		rest.Action = 0
		rest.Database = SqlDataName
		'rest.Devices = Files
		rest.Files = BackupFilePath
		rest.ReplaceDatabase = True
		rest.SQLRestore srv
		If Err.Number>0 Then
			ErrMsg = ErrMsg & "<li>�������ݿ�ʱ��������</li>"
			ErrMsg = ErrMsg & "<li>������룺"
			ErrMsg = ErrMsg & Err.Number & "</li><li><font color=red>"
			'Response.Write Err.Number&"<font color=red><br>"
			ErrMsg = ErrMsg &  Err.Description&"</font></li>"
			FoundErr = True
			Exit Sub
		End If
		Set srv = Nothing
		Set rest = Nothing
		Succeed("<li>SQL���ݿ�ָ��ɹ���</li>")
	Else
		FoundErr = True
		ErrMsg = "����Ŀ¼�²������ı����ļ���"
		Exit Sub
	End If
	Set FSO = Nothing
End Sub
%>



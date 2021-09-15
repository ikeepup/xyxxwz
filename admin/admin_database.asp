<!--#include file="setup.asp"-->
<!--#include file="check.asp"-->
<%
Admin_header
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
Dim bkfolder, bkdbname, fso, fso1
Dim Action
Action = LCase(Request("action"))

Select Case Action
	Case "backupdata" '备份数据
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
	Case "compressdata" '压缩数据
		If Not ChkAdmin("CompressData") Then
			Server.Transfer("showerr.asp")
			Request.End
		End If
		If request("act") = "Compress" Then
			Call CompressDatabase()
		Else
			Call CompressData()
		End If
	Case "restoredata" '恢复数据
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

	Case "spacesize" '系统空间占用
		If Not ChkAdmin("SpaceSize") Then
			Server.Transfer("showerr.asp")
			Request.End
		End If
		Call SpaceSize()

	Case Else
		Errmsg = ErrMsg + "<BR><li>选取相应的操作。"
		ReturnError(ErrMsg)

End Select
If FoundErr = True Then
	ReturnError(ErrMsg)
End If
Admin_footer
SaveLogInfo(AdminName)
CloseConn
'====================系统空间占用=======================
Sub SpaceSize()
	On Error Resume Next
%>
		<table border="0"  cellspacing="1" cellpadding="5" height="1" align=center width="95%" class="tableBorder">			<tr>
  					<th height=25>
  					&nbsp;&nbsp;系统空间占用情况
  					</th>
  				</tr> 	
 				<tr>
 					<td class="TableRow1"> 			
 			<blockquote>
 			<br> 			
 			法规数据占用空间：&nbsp;<img src="images/bar1.gif" width=<%=drawbar("../Database")%> height=10>&nbsp;<%showSpaceinfo("../database")%><br><br>
 			备份数据占用空间：&nbsp;<img src="images/bar1.gif" width=<%=drawbar("databackup")%> height=10>&nbsp;<%showSpaceinfo("databackup")%><br><br>
			后台管理占用空间：&nbsp;<img src="images/bar1.gif" width=<%=drawbar("./")%> height=10>&nbsp;<%showSpaceinfo("./")%><br><br>
 			软件频道占用空间：&nbsp;<img src="images/bar1.gif" width=<%=drawbar("../Soft")%> height=10>&nbsp;<%showSpaceinfo("../Soft")%><br><br>
			文章频道占用空间：&nbsp;<img src="images/bar1.gif" width=<%=drawbar("../Article")%> height=10>&nbsp;<%showSpaceinfo("../Article")%><br><br>
 			模板图片占用空间：&nbsp;<img src="images/bar1.gif" width=<%=drawbar("../skin")%> height=10>&nbsp;<%showSpaceinfo("../skin")%><br><br>
 			图片文件占用空间：&nbsp;<img src="images/bar1.gif" width=<%=drawbar("../images")%> height=10>&nbsp;<%showSpaceinfo("../images")%><br><br>
 			用户文件占用空间：&nbsp;<img src="images/bar1.gif" width=<%=drawbar("../user")%> height=10>&nbsp;<%showSpaceinfo("../user")%><br><br>
			上传文件占用空间：&nbsp;<img src="images/bar1.gif" width=<%=drawbar("../UploadFile")%> height=10>&nbsp;<%showSpaceinfo("../UploadFile")%><br><br>
 			系统占用空间总计：&nbsp;<img src="images/bar2.gif" width=400 height=10> <%showspecialspaceinfo("All")%>
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
  					&nbsp;&nbsp;SQL数据库数据处理说明
  					</th>
  				</tr> 	
 				<tr>
 					<td class="TableRow1"> 			
 			<blockquote>
<B>一、备份数据库</B>
<BR><BR>
1、打开SQL企业管理器，在控制台根目录中依次点开Microsoft SQL Server<BR>
2、SQL Server组-->双击打开你的服务器-->双击打开数据库目录<BR>
3、选择你的数据库名称（如系统数据库enchiasp）-->然后点上面菜单中的工具-->选择备份数据库<BR>
4、备份选项选择完全备份，目的中的备份到如果原来有路径和名称则选中名称点删除，然后点添加，如果原来没有路径和名称则直接选择添加，接着指定路径和文件名，指定后点确定返回备份窗口，接着点确定进行备份
<BR><BR>
<B>二、还原数据库</B><BR><BR>
1、打开SQL企业管理器，在控制台根目录中依次点开Microsoft SQL Server<BR>
2、SQL Server组-->双击打开你的服务器-->点图标栏的新建数据库图标，新建数据库的名字自行取<BR>
3、点击新建好的数据库名称（如系统数据库enchiasp）-->然后点上面菜单中的工具-->选择恢复数据库<BR>
4、在弹出来的窗口中的还原选项中选择从设备-->点选择设备-->点添加-->然后选择你的备份文件名-->添加后点确定返回，这时候设备栏应该出现您刚才选择的数据库备份文件名，备份号默认为1（如果您对同一个文件做过多次备份，可以点击备份号旁边的查看内容，在复选框中选择最新的一次备份后点确定）-->然后点击上方常规旁边的选项按钮<BR>
5、在出现的窗口中选择在现有数据库上强制还原，以及在恢复完成状态中选择使数据库可以继续运行但无法还原其它事务日志的选项。在窗口的中间部位的将数据库文件还原为这里要按照你SQL的安装进行设置（也可以指定自己的目录），逻辑文件名不需要改动，移至物理文件名要根据你所恢复的机器情况做改动，如您的SQL数据库装在D:\Program Files\Microsoft SQL Server\MSSQL\Data，那么就按照您恢复机器的目录进行相关改动改动，并且最后的文件名最好改成您当前的数据库名（如原来是enchicms_data.mdf，现在的数据库是enchiasp，就改成enchiasp_data.mdf），日志和数据文件都要按照这样的方式做相关的改动（日志的文件名是*_log.ldf结尾的），这里的恢复目录您可以自由设置，前提是该目录必须存在（如您可以指定d:\sqldata\enchicms_data.mdf或者d:\sqldata\enchicms_log.ldf），否则恢复将报错<BR>
6、修改完成后，点击下面的确定进行恢复，这时会出现一个进度条，提示恢复的进度，恢复完成后系统会自动提示成功，如中间提示报错，请记录下相关的错误内容并询问对SQL操作比较熟悉的人员，一般的错误无非是目录错误或者文件名重复或者文件名错误或者空间不够或者数据库正在使用中的错误，数据库正在使用的错误您可以尝试关闭所有关于SQL窗口然后重新打开进行恢复操作，如果还提示正在使用的错误可以将SQL服务停止然后重起看看，至于上述其它的错误一般都能按照错误内容做相应改动后即可恢复<BR><BR>

<B>三、收缩数据库</B><BR><BR>
一般情况下，SQL数据库的收缩并不能很大程度上减小数据库大小，其主要作用是收缩日志大小，应当定期进行此操作以免数据库日志过大<BR>
1、设置数据库模式为简单模式：打开SQL企业管理器，在控制台根目录中依次点开Microsoft SQL Server-->SQL Server组-->双击打开你的服务器-->双击打开数据库目录-->选择你的数据库名称（如系统数据库enchiasp）-->然后点击右键选择属性-->选择选项-->在故障还原的模式中选择“简单”，然后按确定保存<BR>
2、在当前数据库上点右键，看所有任务中的收缩数据库，一般里面的默认设置不用调整，直接点确定<BR>
3、<font color=blue>收缩数据库完成后，建议将您的数据库属性重新设置为标准模式，操作方法同第一点，因为日志在一些异常情况下往往是恢复数据库的重要依据</font>
<BR><BR>

<B>四、设定每日自动备份数据库</B><BR><BR>
<font color=red>强烈建议有条件的用户进行此操作！</font><BR>
1、打开企业管理器，在控制台根目录中依次点开Microsoft SQL Server-->SQL Server组-->双击打开你的服务器<BR>
2、然后点上面菜单中的工具-->选择数据库维护计划器<BR>
3、下一步选择要进行自动备份的数据-->下一步更新数据优化信息，这里一般不用做选择-->下一步检查数据完整性，也一般不选择<BR>
4、下一步指定数据库维护计划，默认的是1周备份一次，点击更改选择每天备份后点确定<BR>
5、下一步指定备份的磁盘目录，选择指定目录，如您可以在D盘新建一个目录如：d:\databak，然后在这里选择使用此目录，如果您的数据库比较多最好选择为每个数据库建立子目录，然后选择删除早于多少天前的备份，一般设定4－7天，这看您的具体备份要求，备份文件扩展名一般都是bak就用默认的<BR>
6、下一步指定事务日志备份计划，看您的需要做选择-->下一步要生成的报表，一般不做选择-->下一步维护计划历史记录，最好用默认的选项-->下一步完成<BR>
7、完成后系统很可能会提示Sql Server Agent服务未启动，先点确定完成计划设定，然后找到桌面最右边状态栏中的SQL绿色图标，双击点开，在服务中选择Sql Server Agent，然后点击运行箭头，选上下方的当启动OS时自动启动服务<BR>
8、这个时候数据库计划已经成功的运行了，他将按照您上面的设置进行自动备份
<BR><BR>
修改计划：<BR>
1、打开企业管理器，在控制台根目录中依次点开Microsoft SQL Server-->SQL Server组-->双击打开你的服务器-->管理-->数据库维护计划-->打开后可看到你设定的计划，可以进行修改或者删除操作
<BR><BR>
<B>五、数据的转移（新建数据库或转移服务器）</B><BR><BR>
一般情况下，最好使用备份和还原操作来进行转移数据，在特殊情况下，可以用导入导出的方式进行转移，这里介绍的就是导入导出方式，导入导出方式转移数据一个作用就是可以在收缩数据库无效的情况下用来减小（收缩）数据库的大小，本操作默认为您对SQL的操作有一定的了解，如果对其中的部分操作不理解，可以咨询动网相关人员或者查询网上资料<BR>
1、将原数据库的所有表、存储过程导出成一个SQL文件，导出的时候注意在选项中选择编写索引脚本和编写主键、外键、默认值和检查约束脚本选项<BR>
2、新建数据库，对新建数据库执行第一步中所建立的SQL文件<BR>
3、用SQL的导入导出方式，对新数据库导入原数据库中的所有表内容<BR>
 			</blockquote> 	
 					</td>
 				</tr>
 			</table>
<%
End Sub

'====================恢复数据库=========================

Sub RestoreData()
	If IsSqlDataBase = 1 Then
%>
	<table border="0"  cellspacing="1" cellpadding="5" height="1" align=center width="95%" class="tableBorder">
  				<tr>
  					<th height=25 >
  					&nbsp;&nbsp;恢复SQL数据
  					</th>
  				</tr>
  				<form method="post" action="?action=RestoreData&act=Restore">
  				<tr>
  					<td height=100 class="TableRow1">
  						&nbsp;&nbsp;
						恢复SQL数据库名称：<input type=text size=25 name=SqlDataName value="<%=SqlDatabaseName%>"><BR>&nbsp;&nbsp;
						SQL数据库用户名称：<input type=text size=25 name=SqlUserID value="<%=SqlUsername%>">&nbsp;请输入您的SQL数据库用户名<BR>&nbsp;&nbsp;
						SQL数据库用户密码：<input type=Password size=25 name=SqlUserPass value="<%=SqlPassword%>">&nbsp;请输入您的SQL数据库连接密码<BR>&nbsp;&nbsp;
						SQL数据库服务器名：<input type=text size=25 name=SqlServer value="<%=SqlLocalName%>">&nbsp;连接服务器名（本地用local，外地用IP）<BR>&nbsp;&nbsp;
						备份SQL数据库目录：<input type=text size=25 name=BackupSqlDir value="Databackup">&nbsp;请输入您备份的数据库目录<BR>&nbsp;&nbsp;
						备份SQL数据库名称：<input type=text size=25 name=BackupSqlName  value="$1.bak">&nbsp;请输入您备份的数据库名<BR>
						&nbsp;&nbsp;<input type=submit value="恢复数据库" class=Button><br>
  						-----------------------------------------------------------------------------------------<br>
  						&nbsp;&nbsp;您可以用这个功能来恢复您的法规数据，请确定你的数据库用户有相关权限才能恢复！<br>
  						&nbsp;&nbsp;注意：数据库使用中可能无法恢复	</font>
  					</td>
  				</tr>	
  				</form>
  	</table>
<%
	Else
%>
<table border="0"  cellspacing="1" cellpadding="5" height="1" align=center width="95%" class="tableBorder"		<tr>
	<th height=25 >
   					&nbsp;&nbsp;恢复系统数据 ( 需要FSO支持，FSO相关帮助请看微软网站 )
  					</th>
  				</tr>
				<form method="post" action="?action=RestoreData&act=Restore">
  				
  				<tr>
  					<td height=100 class="TableRow1">
  						&nbsp;&nbsp;备份数据库路径(相对)：<input type=text size=45 name=DBpath value="DataBackup\enchicms_Backup.MDB">&nbsp;&nbsp;<BR>
  						&nbsp;&nbsp;目标数据库路径(相对)：<input type=text size=45 name=backpath value="<%=db%>"><BR>&nbsp;&nbsp;填写您当前使用的数据库路径，如不想覆盖当前文件，可自行命名（注意路径是否正确），然后修改conn.asp文件，如果目标文件名和当前使用数据库名一致的话，不需修改conn.asp文件<BR>
						&nbsp;&nbsp;<input type=submit value="恢复数据库" onclick="{if(confirm('您确定要恢复数据库吗?')){return true;}return false;}" class=Button> <br>
  						-----------------------------------------------------------------------------------------<br>
  						&nbsp;&nbsp;在上面填写本程序的数据库路径全名，本程序的默认备份数据库文件为DataBackup\enchicms_Backup.MDB，请按照您的备份文件自行修改。<br>
  						&nbsp;&nbsp;您可以用这个功能来备份您的法规数据，以保证您的数据安全！<br>
  						&nbsp;&nbsp;注意：所有路径都是相对与程序空间根目录的相对路径</font>
  					</td>
  				</tr>	
  				</form>
  			</table>
<%
	End If
End Sub
'====================备份数据库=========================
Sub BackupData()
	If IsSqlDataBase = 1 Then
%>
	<table border="0"  cellspacing="1" cellpadding="5" height="1" align=center width="95%" class="tableBorder">
  				<tr>
  					<th height=25 >
  					&nbsp;&nbsp;备份SQL数据
  					</th>
  				</tr>
  				<form method="post" action="?action=BackupData&act=Backup">
  				<tr>
  					<td height=100 class="TableRow1">
  						&nbsp;&nbsp;
						当前SQL数据库名称：<input type=text size=25 name=SqlDataName value="<%=SqlDatabaseName%>"><BR>&nbsp;&nbsp;
						SQL数据库用户名称：<input type=text size=25 name=SqlUserID value="<%=SqlUsername%>">&nbsp;请输入您的SQL数据库用户名<BR>&nbsp;&nbsp;
						SQL数据库用户密码：<input type=Password size=25 name=SqlUserPass value="<%=SqlPassword%>">&nbsp;请输入您的SQL数据库连接密码<BR>&nbsp;&nbsp;
						SQL数据库服务器名：<input type=text size=25 name=SqlServer value="<%=SqlLocalName%>">&nbsp;连接服务器名（本地用local，外地用IP）<BR>&nbsp;&nbsp;
						备份SQL数据库目录：<input type=text size=25 name=BackupSqlDir value="Databackup">&nbsp;如目录不存在，程序将自动创建<BR>&nbsp;&nbsp;
						备份SQL数据库名称：<input type=text size=25 name=BackupSqlName  value="$1.bak">&nbsp;如使用默认的($1)备份名,系统将自动按日期时间创建备份名称<BR>
						&nbsp;&nbsp;<input type=submit value="备份数据库" class=Button><br>
  						-----------------------------------------------------------------------------------------<br>
  						&nbsp;&nbsp;您可以用这个功能来备份您的法规数据，以保证您的数据安全！<br>
  						&nbsp;&nbsp;注意：所有路径都是相对与程序空间管理目录的相对路径				</font>
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
  					&nbsp;&nbsp;备份系统数据 ( 需要FSO支持，FSO相关帮助请看微软网站 )
  					</th>
  				</tr>
  				<form method="post" action="?action=BackupData&act=Backup">
  				<tr>
  					<td height=100 class="TableRow1">
  						&nbsp;&nbsp;
						当前数据库路径(相对路径)：<input type=text size=45 name=DBpath size=45 value="<%=db%>"><BR>&nbsp;&nbsp;
						备份数据库目录(相对路径)：<input type=text size=20 name=bkfolder size=45 value=Databackup>&nbsp;如目录不存在，程序将自动创建<BR>&nbsp;&nbsp;
						备份数据库名称(填写名称)：<input type=text size=20 name=bkDBname size=45 value=enchicms_Backup.MDB>&nbsp;如备份目录有该文件，将覆盖，如没有，将自动创建<BR>
						&nbsp;&nbsp;<input type=submit value="备份数据库" class=Button><br>
  						-----------------------------------------------------------------------------------------<br>
  						&nbsp;&nbsp;在上面填写本程序的数据库路径全名，本程序的默认数据库文件为<%=db%>，<B>请一定不能用默认名称命名备份数据库</B><br>
  						&nbsp;&nbsp;您可以用这个功能来备份您的法规数据，以保证您的数据安全！<br>
  						&nbsp;&nbsp;注意：所有路径都是相对与程序空间管理目录的相对路径				</font>
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
		Succeed("备份数据库成功，您备份的数据库路径为" &bkfolder& "\"& bkdbname)
	Else
		FoundErr = True
		ErrMsg = "找不到您所需要备份的文件。"
		Exit Sub
	End If
End Sub
Sub RestoreDatabase()
	Dim backpath,Dbpath
	Dbpath = request.Form("Dbpath")
	backpath = request.Form("backpath")
	If dbpath = "" Then
		FoundErr = True
		ErrMsg = "请输入您要恢复成的数据库全名"
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
		Succeed("成功恢复数据！")
	Else
		FoundErr = True
		ErrMsg = "备份目录下并无您的备份文件！"
		Exit Sub
	End If
End Sub
'------------------检查某一目录是否存在-------------------
Function CheckDir(FolderPath)
	folderpath = Server.MapPath(".")&"\"&folderpath
	Set fso1 = CreateObject("Scripting.FileSystemObject")
	If fso1.FolderExists(FolderPath) Then
		'存在
		CheckDir = True
	Else
		'不存在
		CheckDir = False
	End If
	Set fso1 = Nothing
End Function
'-------------根据指定名称生成目录-----------------------
Function MakeNewsDir(foldername)
	Dim f
	Set fso1 = CreateObject("Scripting.FileSystemObject")
	Set f = fso1.CreateFolder(foldername)
	MakeNewsDir = True
	Set fso1 = Nothing
End Function
'====================压缩数据库 =========================
Sub CompressData()

	If IsSqlDataBase =1 Then
		SQLUserReadme()
		Exit Sub
	End If
%>
<table border="0"  cellspacing="1" cellpadding="5" height="1" align=center width="95%" class="tableBorder">
<tr>
<th height=25 >
&nbsp;&nbsp;压缩数据库 ( 需要FSO支持，FSO相关帮助请看微软网站 )
</th>
<form action="?action=CompressData&act=Compress" method="post">
<tr>
<td class="TableRow1" height=25><b>注意：</b><br>输入数据库所在相对路径,并且输入数据库名称（正在使用中数据库不能压缩，请选择备份数据库进行压缩操作） </td>
</tr>
<tr>
<td class="TableRow1">压缩数据库：<input type="text" name="dbpath" size=45 value=<%=db%>>&nbsp;
<input type="submit" value="开始压缩" class=Button></td>
</tr>
<tr>
<td class="TableRow1"><input type="checkbox" name="boolIs97" value="True">如果使用 Access 97 数据库请选择
(默认为 Access 2000 数据库)<br><br></td>
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
		ErrMsg = "请输入要压缩的数据库路径！"
		Exit Sub
	End If
End Sub
'=====================压缩参数=========================
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
		Succeed("你的数据库, " & dbpath & ", 已经压缩成功!")
	Else
		ReturnError("数据库名称或路径不正确. 请重试!")
	End If

End Function
'=====================系统空间参数=========================
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
		ErrMsg = ErrMsg & "<li>请输入SQL数据库名！</li>"
	End If
	If Trim(Request.Form("SqlUserPass")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg & "<li>请输入SQL数据库用户密码！</li>"
	End If
	If Trim(Request.Form("SqlUserID")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg & "<li>请输入SQL数据库用户名称！</li>"
	End If
	If Trim(Request.Form("SqlServer")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg & "<li>请输入SQL数据库连接名（本地用local，外地用IP）！</li>"
	End If
	If Trim(Request.Form("BackupSqlName")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg & "<li>请输入SQL数据库备份名称！</li>"
	End If
	If Trim(Request.Form("BackupSqlDir")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg & "<li>请输入SQL数据库备份目录！</li>"
	End If
End Sub
'====================备份SQL数据库=========================
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
	SqlLoginTimeout = 20 '登陆超时
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
	Succeed("<li>SQL数据库备份成功！</li>")
End Sub
'====================恢复SQL数据库=========================
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
	SqlLoginTimeout = 20 '登陆超时
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
			ErrMsg = ErrMsg & "<li>备份数据库时发生错误！</li>"
			ErrMsg = ErrMsg & "<li>错误代码："
			ErrMsg = ErrMsg & Err.Number & "</li><li><font color=red>"
			'Response.Write Err.Number&"<font color=red><br>"
			ErrMsg = ErrMsg &  Err.Description&"</font></li>"
			FoundErr = True
			Exit Sub
		End If
		Set srv = Nothing
		Set rest = Nothing
		Succeed("<li>SQL数据库恢复成功！</li>")
	Else
		FoundErr = True
		ErrMsg = "备份目录下并无您的备份文件！"
		Exit Sub
	End If
	Set FSO = Nothing
End Sub
%>



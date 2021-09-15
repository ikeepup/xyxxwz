<!--#include file="md5.asp"-->
<%
Const IsDeBug = 1
Class enchiaspMain_Cls
	
	Public membername, memberpass, membergrade, membergroup, memberid
	Public memberclass, menbernickname, Cookies_Name, CheckPassword

	Public SiteName, SiteUrl, MasterMail, keywords, Copyright
	Public InstallDir, IndexName, IstopSite, StopReadme, IsCloseMail
	Public SendMailType, MailFrom, MailServer, MailUserName, MailPassword, MailInformPass, ChkSameMail
	Public CheckUserReg, AdminCheckReg, AddUserPoint, SendRegMessage, FullContQuery, ActionTime
	Public IsRunTime, UploadClass, UploadFileSize, UploadFileType, ContentKeyword, PreviewSetting
	Public StopApplyLink, FSO_ScriptName, InitTitleColor, StopBankPay
	Public ChinaeBank, VersionID, Badwords, Badwordr, serialcode, passedcode
	public usefengmian,fengmianname,fengmiannametop,fengmiannameleft
	public ercilogin,mypass,mypasskey
	public url,urlreg,urldate,urlflag,urlman
	public kkstr

	Public ChannelName, ChannelDir, StopChannel, ChannelType
	Public modules, ChannelSkin, HtmlPath, HtmlForm, HtmlPrefix
	Public IsCreateHtml, HtmlExtName, StopUpload, MaxFileSize, UpFileType
	Public IsAuditing, AppearGrade, ModuleName, BindDomain, DomainName
	Public PostGrade, LeastString, MaxString, PaginalNum, LeastHotHist, Channel_Setting
	Public ChannelSetting,ChannelData,ChannelPath
	Public ChannelModule,ChannelHtmlPath,ChannelHtmlForm,ChannelUseHtml,ChannelHtmlExt,ChannelPrefix
	
	Public ThisEdition, CopyrightStr, Version, Values, startime
	Public SqlQueryNum, GetUserip, CacheName, Reloadtime

	Public ScriptName, Admin_Page, skinid, SkinPath, HtmlCss, HtmlTop, HtmlFoot, HtmlContent, sHtmlContent
	Private Main_Style, Main_Setting, MainStyle, Html_Setting
	Private LocalCacheName, Cache_Data
	Private CacheChannel, CacheData

	Private arrGroupSetting, blnGroupSetting, binUserLong
	
	'图片变换的路径，说明等
	public tupianhuanpic,tupianhuanlink,tupianhuantext

	public tupianhuanpic2,tupianhuanlink2,tupianhuantext2

	public tupianhuanpic3,tupianhuanlink3,tupianhuantext3

	public tupianhuanpic4,tupianhuanlink4,tupianhuantext4

	public tupianhuanpic5,tupianhuanlink5,tupianhuantext5

	public tupianhuanpic6,tupianhuanlink6,tupianhuantext6

	public dibutupian
	'首页视频动画
	public vodpath
	
	Private Sub Class_Initialize()
		On Error Resume Next
		Reloadtime = 28800
		SqlQueryNum = 0
		'--缓存名称
		CacheName = "enchiasp"
		Cookies_Name = "enchiasp_net"
		binUserLong = False
		blnGroupSetting = False
		GetUserip = Request.ServerVariables("HTTP_X_FORWARDED_FOR")
		If Len(GetUserip) = 0 Then GetUserip = Request.ServerVariables("REMOTE_ADDR")
		GetUserip = CheckStr(GetUserip)
		membername = CheckStr(Request.Cookies(Cookies_Name)("username"))
		memberpass = CheckStr(Request.Cookies(Cookies_Name)("password"))
		menbernickname = CheckStr(Request.Cookies(Cookies_Name)("nickname"))
		membergrade = ChkNumeric(Request.Cookies(Cookies_Name)("UserGrade"))
		membergroup = CheckStr(Request.Cookies(Cookies_Name)("UserGroup"))
		memberclass = ChkNumeric(Request.Cookies(Cookies_Name)("UserClass"))
		memberid = ChkNumeric(Request.Cookies(Cookies_Name)("userid"))
		CheckPassword = CheckStr(Request.Cookies(Cookies_Name)("CheckPassword"))
		Dim tmpstr, i
		tmpstr = Request.ServerVariables("PATH_INFO")
		tmpstr = Split(tmpstr, "/")
		i = UBound(tmpstr)
		ScriptName = LCase(tmpstr(i))
		Admin_Page = False
		If InStr(ScriptName, "showerr") > 0 Or InStr(ScriptName, "login") > 0 Or InStr(ScriptName, "admin_") > 0 Then Admin_Page = True
	End Sub

	Private Sub Class_Terminate()
		If IsObject(Conn) Then Conn.Close : Set Conn = Nothing
	End Sub

	'===================服务器缓存部分函数开始===================
	Public Property Let Name(ByVal vNewValue)
		LocalCacheName = LCase(vNewValue)
		Cache_Data = Application(CacheName & "_" & LocalCacheName)
	End Property
	Public Property Let Value(ByVal vNewValue)
		If LocalCacheName <> "" Then
			ReDim Cache_Data(2)
			Cache_Data(0) = vNewValue
			Cache_Data(1) = Now()
			Application.Lock
			Application(CacheName & "_" & LocalCacheName) = Cache_Data
			Application.UnLock
		Else
			Err.Raise vbObjectError + 1, "enchiaspCacheServer", " please change the CacheName."
		End If
	End Property
	Public Property Get Value()
		If LocalCacheName <> "" Then
			If IsArray(Cache_Data) Then
				Value = Cache_Data(0)
			Else
				'Err.Raise vbObjectError + 1, "enchiaspCacheServer", " The Cache_Data("&LocalCacheName&") Is Empty."
			End If
		Else
			Err.Raise vbObjectError + 1, "enchiaspCacheServer", " please change the CacheName."
		End If
	End Property
	Public Function ObjIsEmpty()
		ObjIsEmpty = True
		If Not IsArray(Cache_Data) Then Exit Function
		If Not IsDate(Cache_Data(1)) Then Exit Function
		If DateDiff("s", CDate(Cache_Data(1)), Now()) < (60 * Reloadtime) Then ObjIsEmpty = False
	End Function
	Public Sub DelCahe(MyCaheName)
		Application.Lock
		Application.Contents.Remove (CacheName & "_" & MyCaheName)
		Application.UnLock
	End Sub
	Public Sub DelCache(MyCaheName)
		Application.Lock
		Application.Contents.Remove ("myenchiasp_" & MyCaheName)
		Application.UnLock
	End Sub
	'===================服务器缓存部分函数结束===================
	
	Public Function ChkBoolean(ByVal Values)
		If TypeName(Values) = "Boolean" Or IsNumeric(Values) Or LCase(Values) = "false" Or LCase(Values) = "true" Then
			ChkBoolean = CBool(Values)
		Else
			ChkBoolean = False
		End If
	End Function

	Public Function CheckNumeric(ByVal CHECK_ID)
		If CHECK_ID <> "" And IsNumeric(CHECK_ID) Then
			CHECK_ID = CCur(CHECK_ID)
		Else
			CHECK_ID = 0
		End If
		CheckNumeric = CHECK_ID
	End Function

	Public Function ChkNumeric(ByVal CHECK_ID)
		If CHECK_ID <> "" And IsNumeric(CHECK_ID) Then
			CHECK_ID = CLng(CHECK_ID)
			If CHECK_ID < 0 Then CHECK_ID = 0
		Else
			CHECK_ID = 0
		End If
		ChkNumeric = CHECK_ID
	End Function

	Public Function CheckStr(ByVal str)
		If IsNull(str) Then
			CheckStr = ""
			Exit Function
		End If
		str = Replace(str, Chr(0), "")
		CheckStr = Replace(str, "'", "''")
	End Function
	'================================================
	'过程名：CheckNull
	'作  用：是否有效值
	'================================================
	Public Function CheckNull(ByVal sValue)
		On Error Resume Next
		If IsNull(sValue) Then
			CheckNull = False
			Exit Function
		End If
		If Trim(sValue) <> "" And LCase(Trim(sValue)) <> "http://" Then
			CheckNull = True
		Else
			CheckNull = False
		End If
	End Function
	Public Function ChkNull(ByVal str)
		On Error Resume Next
		If IsNull(str) Then
			ChkNull = ""
			Exit Function
		End If
		If Trim(str) <> "" And LCase(Trim(str)) <> "http://" Then
			ChkNull = Trim(str)
		Else
			ChkNull = ""
		End If
	End Function
	'=============================================================
	'函数名：ChkFormStr
	'作  用：过滤表单字符
	'参  数：str   ----原字符串
	'返回值：过滤后的字符串
	'=============================================================
	Public Function ChkFormStr(ByVal str)
		Dim fString
		fString = str
		If IsNull(fString) Then
			ChkFormStr = ""
			Exit Function
		End If
		fString = Replace(fString, "'", "&#39;")
		fString = Replace(fString, Chr(34), "&quot;")
		fString = Replace(fString, Chr(13), "")
		fString = Replace(fString, Chr(10), "")
		fString = Replace(fString, Chr(9), "")
		fString = Replace(fString, ">", "&gt;")
		fString = Replace(fString, "<", "&lt;")
		fString = Replace(fString, "%", "％")
		ChkFormStr = Trim(JAPEncode(fString))
	End Function
	'=============================================================
	'函数作用：过滤SQL非法字符
	'=============================================================
	Public Function CheckRequest(ByVal str,ByVal strLen)
		On Error Resume Next
		str = Trim(str)
		str = Replace(str, Chr(0), "")
		str = Replace(str, "'", "")
		str = Replace(str, "%", "")
		str = Replace(str, "^", "")
		str = Replace(str, ";", "")
		str = Replace(str, "*", "")
		str = Replace(str, "<", "")
		str = Replace(str, ">", "")
		str = Replace(str, "|", "")
		str = Replace(str, "and", "")
		str = Replace(str, "chr", "")
		
		If Len(str) > 0 And strLen > 0 Then
			str = Left(str, strLen)
		End If
		CheckRequest = str
	End Function
	'-- 移除有害字符
	Public Function RemoveBadCharacters(ByVal strTemp)
		Dim re
		On Error Resume Next
		Set re = New RegExp
		re.Pattern = "[^\s\w]"
		re.Global = True
		RemoveBadCharacters = re.Replace(strTemp, "")
		Set re = Nothing
	End Function
	'-- 去掉HTML标记
	Public Function RemoveHtml(ByVal Textstr)
		Dim Str,re
		Str = Textstr
		On Error Resume Next
		Set re = New RegExp
		re.IgnoreCase = True
		re.Global = True
		re.Pattern = "<(.[^>]*)>"
		Str = re.Replace(Str, "")
		Set re = Nothing
		RemoveHtml=Str
	End Function
	'-- 数据库连接
	Public Function Execute(Command)
		If Not IsObject(Conn) Then ConnectionDatabase		
		If IsDeBug = 0 Then 
			On Error Resume Next
			Set Execute = Conn.Execute(Command)
			If Err Then
				err.Clear
				Set Conn = Nothing
				Response.Write "查询数据的时候发现错误，请检查您的查询代码是否正确。<br /><li>"
				Response.Write Command
				Response.End
			End If
		Else
			Set Execute = Conn.Execute(Command)
		End If	
		SqlQueryNum = SqlQueryNum+1
	End Function
	
	Public Sub ReadConfig()
		On Error Resume Next
		Name = "Config"
		If ObjIsEmpty() Then ReloadConfig
		CacheData = Value
		'第一次起用系统或者重启IIS的时候加载缓存
		Name = "Date"
		If ObjIsEmpty() Then
			Value = Date
		Else
			If CStr(Value) <> CStr(Date) Then
				Name = "Config"
				Call ReloadConfig
				CacheData = Value
			End If
		End If
		SiteName = CacheData(1, 0): SiteUrl = CacheData(2, 0): MasterMail = CacheData(3, 0): keywords = CacheData(4, 0): Copyright = CacheData(5, 0): InstallDir = CacheData(6, 0)
		IndexName = CacheData(7, 0): IstopSite = CacheData(8, 0): StopReadme = CacheData(9, 0): IsCloseMail = CacheData(10, 0): SendMailType = CacheData(11, 0): MailFrom = CacheData(12, 0)
		MailServer = CacheData(13, 0): MailUserName = CacheData(14, 0): MailPassword = CacheData(15, 0): CheckUserReg = CacheData(16, 0): AdminCheckReg = CacheData(17, 0): MailInformPass = CacheData(18, 0)
		ChkSameMail = CacheData(19, 0): AddUserPoint = CacheData(20, 0): SendRegMessage = CacheData(21, 0): FullContQuery = CacheData(22, 0): ActionTime = CacheData(23, 0): IsRunTime = CacheData(24, 0)
		UploadClass = CacheData(25, 0): UploadFileSize = CacheData(26, 0): UploadFileType = CacheData(27, 0): ContentKeyword = CacheData(28, 0): StopApplyLink = CacheData(29, 0): FSO_ScriptName = CacheData(30, 0)
		InitTitleColor = CacheData(31, 0): StopBankPay = CacheData(32, 0): ChinaeBank = CacheData(33, 0): VersionID = CacheData(34, 0): Badwords = CacheData(35, 0): Badwordr = CacheData(36, 0)
		serialcode = CacheData(37, 0): passedcode = CacheData(38, 0) : PreviewSetting = CacheData(39, 0)
		ThisEdition = "企业版 (Enterprise Edition)"
		Version = "Powered by：<a href=""http://www.enchi.com.cn"" target=""_blank""  class=""navmenu"">enchicms SiteManageSystem Version 3.0.0</a>"
		CopyrightStr = "<!--" & vbCrLf
		CopyrightStr = CopyrightStr & "┌─────────enchicms─────────┐" & vbCrLf
		CopyrightStr = CopyrightStr & "│enchicms  Version 3.0.0                     │" & vbCrLf
		CopyrightStr = CopyrightStr & "│版权所有: 恩池软件                          │" & vbCrLf
		CopyrightStr = CopyrightStr & "│官方主页: http://www.enchi.com.cn           │" & vbCrLf
		CopyrightStr = CopyrightStr & "│论坛地址: http://www.enchi.com.cn           │" & vbCrLf
		CopyrightStr = CopyrightStr & "│E-Mail:   liuyunfan@163.com  QQ: 21556923   │" & vbCrLf
		CopyrightStr = CopyrightStr & "└──────────────────────┘" & vbCrLf
		CopyrightStr = CopyrightStr & "-->" & vbCrLf
		usefengmian=CacheData(40, 0)
		fengmianname=CacheData(41, 0)
		fengmiannametop=CacheData(42, 0)

		fengmiannameleft=CacheData(43, 0)
		ercilogin=CacheData(44, 0)
		mypass=CacheData(45, 0)
		mypasskey=CacheData(46, 0)
		url=CacheData(47, 0)
		urlflag=CacheData(48, 0)
		urlman=CacheData(49, 0)
		urldate=CacheData(50, 0)
		urlreg=CacheData(51, 0)
		kkstr=CacheData(52, 0)
		'wordhelp()
		If CInt(IstopSite) = 1 And Not Admin_Page Then Response.Redirect ("" & SiteUrl & InstallDir & "showerr.asp?action=stop")
		gettupian()
gettupian2()
gettupian3()
gettupian4()
gettupian5()
gettupian6()		
	End Sub
	

public sub gettupian2
	'以下内容用来图片变换时使用
	dim temp1,temp2,temp3,temp4,temp5		
		temp1=""
		temp2=""
		temp3=""
		Dim Rs
		Set Rs = Execute("Select * from eccms_tupian2 order by id")
			if Rs.Bof and Rs.Eof then
				temp1=""
				temp2=""
				temp3=""
			else
				do while not rs.eof
					if rs("pic")<>"" then
						temp1=temp1 & rs("pic") 
						temp2=temp2 & rs("picurl") 
						temp3=temp3 & rs("pictext") 
					end if
					rs.movenext
					if not  rs.eof then
						if rs("pic")<>"" then
							temp1=temp1 & "|"
							temp2=temp2 & "|"
							temp3=temp3 & "|"
						end if		
					end if
				loop
			end if
			rs.close
			Set Rs = Nothing	
		tupianhuanpic2=temp1
		tupianhuanlink2=temp2
		tupianhuantext2=temp3			
	end sub

public sub gettupian3
	'以下内容用来图片变换时使用
	dim temp1,temp2,temp3,temp4,temp5		
		temp1=""
		temp2=""
		temp3=""
		Dim Rs
		Set Rs = Execute("Select * from eccms_tupian3 order by id")
			if Rs.Bof and Rs.Eof then
				temp1=""
				temp2=""
				temp3=""
			else
				do while not rs.eof
					if rs("pic")<>"" then
						temp1=temp1 & rs("pic") 
						temp2=temp2 & rs("picurl") 
						temp3=temp3 & rs("pictext") 
					end if
					rs.movenext
					if not  rs.eof then
						if rs("pic")<>"" then
							temp1=temp1 & "|"
							temp2=temp2 & "|"
							temp3=temp3 & "|"
						end if		
					end if
				loop
			end if
			rs.close
			Set Rs = Nothing	
		tupianhuanpic3=temp1
		tupianhuanlink3=temp2
		tupianhuantext3=temp3			
	end sub


public sub gettupian4
	'以下内容用来图片变换时使用
	dim temp1,temp2,temp3,temp4,temp5		
		temp1=""
		temp2=""
		temp3=""
		Dim Rs
		Set Rs = Execute("Select * from eccms_tupian4 order by id")
			if Rs.Bof and Rs.Eof then
				temp1=""
				temp2=""
				temp3=""
			else
				do while not rs.eof
					if rs("pic")<>"" then
						temp1=temp1 & rs("pic") 
						temp2=temp2 & rs("picurl") 
						temp3=temp3 & rs("pictext") 
					end if
					rs.movenext
					if not  rs.eof then
						if rs("pic")<>"" then
							temp1=temp1 & "|"
							temp2=temp2 & "|"
							temp3=temp3 & "|"
						end if		
					end if
				loop
			end if
			rs.close
			Set Rs = Nothing	
		tupianhuanpic4=temp1
		tupianhuanlink4=temp2
		tupianhuantext4=temp3			
	end sub



public sub gettupian5
	'以下内容用来图片变换时使用
	dim temp1,temp2,temp3,temp4,temp5		
		temp1=""
		temp2=""
		temp3=""
		Dim Rs
		Set Rs = Execute("Select * from eccms_tupian5 order by id")
			if Rs.Bof and Rs.Eof then
				temp1=""
				temp2=""
				temp3=""
			else
				do while not rs.eof
					if rs("pic")<>"" then
						temp1=temp1 & rs("pic") 
						temp2=temp2 & rs("picurl") 
						temp3=temp3 & rs("pictext") 
					end if
					rs.movenext
					if not  rs.eof then
						if rs("pic")<>"" then
							temp1=temp1 & "|"
							temp2=temp2 & "|"
							temp3=temp3 & "|"
						end if		
					end if
				loop
			end if
			rs.close
			Set Rs = Nothing	
		tupianhuanpic5=temp1
		tupianhuanlink5=temp2
		tupianhuantext5=temp3			
	end sub



public sub gettupian6
	'以下内容用来图片变换时使用
	dim temp1,temp2,temp3,temp4,temp5		
		temp1=""
		temp2=""
		temp3=""
		Dim Rs
		Set Rs = Execute("Select * from eccms_tupian6 order by id")
			if Rs.Bof and Rs.Eof then
				temp1=""
				temp2=""
				temp3=""
			else
				do while not rs.eof
					if rs("pic")<>"" then
						temp1=temp1 & rs("pic") 
						temp2=temp2 & rs("picurl") 
						temp3=temp3 & rs("pictext") 
					end if
					rs.movenext
					if not  rs.eof then
						if rs("pic")<>"" then
							temp1=temp1 & "|"
							temp2=temp2 & "|"
							temp3=temp3 & "|"
						end if		
					end if
				loop
			end if
			rs.close
			Set Rs = Nothing	
		tupianhuanpic6=temp1
		tupianhuanlink6=temp2
		tupianhuantext6=temp3			
	end sub



	public sub gettupian
	'以下内容用来图片变换时使用
	dim temp1,temp2,temp3,temp4,temp5		
		temp1=""
		temp2=""
		temp3=""
		Dim Rs
		Set Rs = Execute("Select * from eccms_tupian order by id")
			if Rs.Bof and Rs.Eof then
				temp1=""
				temp2=""
				temp3=""
			else
				do while not rs.eof
					if rs("pic")<>"" then
						temp1=temp1 & rs("pic") 
						temp2=temp2 & rs("picurl") 
						temp3=temp3 & rs("pictext") 
					end if
					rs.movenext
					if not  rs.eof then
						if rs("pic")<>"" then
							temp1=temp1 & "|"
							temp2=temp2 & "|"
							temp3=temp3 & "|"
						end if		
					end if
				loop
			end if
			rs.close
			Set Rs = Nothing	
		tupianhuanpic=temp1
		tupianhuanlink=temp2
		tupianhuantext=temp3
		
		
		
		temp4=""		
		Set rs =Execute("Select * from eccms_dibu order by id")	
			if Rs.Bof and Rs.Eof then
			else
				do while not rs.eof			
					if rs("pic")<>"" then
						temp4=temp4&" <li>" 
						temp4=temp4&"<a href='"& rs("picurl") &"'>"
						temp4=temp4&"<img src='"& rs("pic") &"'></a>"
					end if
					rs.movenext
					loop
			end if
			rs.movefirst
			temp4=temp4&" </li>"
					

												
             rs.close
			Set Rs = Nothing                          
		dibutupian=temp4										
	
	
		Set rs =Execute("Select * from eccms_vod")	
			if Rs.Bof and Rs.Eof then
			vodpath="" 
			else
				vodpath=rs("path") 
			end if								
             rs.close
			Set Rs = Nothing 
			      

			
	end sub
	
	
	Public Sub ReloadConfig()
		Dim SQL, Rs
		On Error Resume Next
		SQL = "SELECT * from [ECCMS_Config] "
		Set Rs = Execute(SQL)
		Value = Rs.GetRows(1)
		Set Rs = Nothing
	End Sub
	'=============================================================
	'过程名：ReloadChannel
	'作  用：再装频道设置
	'参  数：ChannelID   ----频道ID
	'=============================================================
	Private Sub ReloadChannel(ChannelID)
		Dim SQL, Rs
		On Error Resume Next
		SQL = "SELECT ChannelID,ChannelName,ChannelDir,StopChannel,ChannelType,modules,ModuleName,BindDomain,DomainName,ChannelSkin,HtmlPath,HtmlForm,IsCreateHtml,HtmlExtName,HtmlPrefix,StopUpload,MaxFileSize,UpFileType,IsAuditing,AppearGrade,PostGrade,LeastString,MaxString,PaginalNum,LeastHotHist,Channel_Setting from ECCMS_Channel where ChannelType <= 1 And ChannelID = " & CLng(ChannelID)
		Set Rs = Execute(SQL)
		If Rs.BOF And Rs.EOF Then
			Response.Write "错误的频道参数！"
			Exit Sub
		End If
		Value = Rs.GetRows(1)
		Set Rs = Nothing
	End Sub
	
	
	
	
	
	
	
	
	
	'=============================================================
	'过程名：ReadChannel
	'作  用：读取频道设置
	'参  数：ChannelID   ----频道ID
	'=============================================================
	Public Sub ReadChannel(ChannelID)
		On Error Resume Next
		If Not IsNumeric(ChannelID) Then ChannelID = 1
		ChannelID = Clng(ChannelID)
		Name = "Channel" & ChannelID
		If ObjIsEmpty() Then Call ReloadChannel(ChannelID)
		CacheChannel = Value
		If CLng(CacheChannel(0, 0)) <> ChannelID Then
			Call ReloadChannel(ChannelID)
			CacheChannel = Value
		End If
		ChannelName = CacheChannel(1, 0): ChannelDir = CacheChannel(2, 0): StopChannel = CacheChannel(3, 0): ChannelType = CacheChannel(4, 0): modules = CacheChannel(5, 0): ModuleName = CacheChannel(6, 0): BindDomain = CacheChannel(7, 0): DomainName = CacheChannel(8, 0): ChannelSkin = CacheChannel(9, 0): HtmlPath = CacheChannel(10, 0)
		HtmlForm = CacheChannel(11, 0): IsCreateHtml = CacheChannel(12, 0): HtmlExtName = CacheChannel(13, 0): HtmlPrefix = CacheChannel(14, 0): StopUpload = CacheChannel(15, 0): MaxFileSize = CacheChannel(16, 0): UpFileType = CacheChannel(17, 0): IsAuditing = CacheChannel(18, 0): AppearGrade = CacheChannel(19, 0)
		PostGrade = CacheChannel(20, 0): LeastString = CacheChannel(21, 0): MaxString = CacheChannel(22, 0): PaginalNum = CacheChannel(23, 0): LeastHotHist = CacheChannel(24, 0): Channel_Setting = CacheChannel(25, 0)
		If CInt(StopChannel) = 1 And Not Admin_Page Then Response.Redirect (InstallDir & "showerr.asp?action=ChanStop")
	End Sub
	
	Public Sub LoadChannel(chanid)
		On Error Resume Next
		Dim Rs,SQL,tmpdata
		chanid = CLng(chanid)
		Name = "MyChannel" & chanid
		If ObjIsEmpty() Then
			SQL = "SELECT ChannelName,ChannelDir,ModuleName,HtmlPath,HtmlForm,IsCreateHtml,HtmlExtName,HtmlPrefix,StopUpload,LeastString,MaxString,LeastHotHist FROM ECCMS_Channel WHERE ChannelType<=1 And ChannelID= " & Clng(chanid)
			Set Rs = Execute(SQL)
			tmpdata = Rs.GetString(, , "|||", "@@@", "")
			tmpdata = Left(tmpdata, Len(tmpdata) - 3)
			Set Rs = Nothing
			Value = tmpdata
		End If
		
		ChannelData = Split(Value, "|||")
		ChannelPath = InstallDir & ChannelData(1)
		ChannelModule = ChannelData(2)
		ChannelHtmlPath = ChannelData(3)
		ChannelHtmlForm = ChannelData(4)
		ChannelUseHtml = ChannelData(5)
		ChannelHtmlExt = ChannelData(6)
		ChannelPrefix = ChannelData(7)
		
	End Sub
	'=============================================================
	'过程名：LoadTemplates
	'作  用：载入模板
	'参  数：Page_Mark   ----StyleID
	'=============================================================
	Public Sub LoadTemplates(ChannelID, pageid, StyleID)
		Dim rstmp, TempSkinID
		On Error Resume Next
		ChannelID = CLng(ChannelID)
		pageid = CInt(pageid)
		Name = "DefaultSkinID"
		If ObjIsEmpty() Then
			Set rstmp = Execute("SELECT skinid from [ECCMS_Template] where pageid = 0 And isDefault = 1")
			Value = rstmp(0)
			Set rstmp = Nothing
		End If
		TempSkinID = Value
		If StyleID = 0 Or StyleID = "" Then
			skinid = TempSkinID
		Else
			Set rstmp = Execute("SELECT skinid from [ECCMS_Template] where pageid = 0 And skinid = " & StyleID)
			If Not rstmp.EOF Then
				skinid = rstmp(0)
			Else
				skinid = TempSkinID
			End If
			Set rstmp = Nothing
		End If
		skinid = CLng(skinid)
		Name = "MainStyle" & skinid
		If ObjIsEmpty() Then TemplatesMainCache (skinid)
		Main_Style = Value
		SkinPath = Main_Style(0, 0)
		Main_Setting = Split(Main_Style(2, 0), "|||")
		MainStyle = Main_Style(1, 0)
		'MainStyle = Replace(MainStyle, "{$InstallDir}", ReadInstallDir(BindDomain))
		MainStyle = Replace(MainStyle, "{$SkinPath}", SkinPath)
		MainStyle = Split(MainStyle, "|||")
		HtmlCss = MainStyle(0)
		HtmlTop = MainStyle(1)
		HtmlFoot = MainStyle(2)
		If pageid <> 0 Then
			Name = "Templates" & ChannelID & skinid & pageid
			If ObjIsEmpty() Then
				TemplatesToCache ChannelID, pageid
			End If
			ByValue = Value
		End If
	End Sub
	
	
	
	Private Sub TemplatesToCache(ChannelID, pageid)
		On Error Resume Next
		Dim Rs, SQL, rstmp
		SQL = "SELECT skinid,page_content,page_setting FROM [ECCMS_Template] WHERE ChannelID = " & ChannelID & " And skinid = " & skinid & " And pageid = " & pageid
		Set Rs = Execute(SQL)
		If Not Rs.EOF Then
			Value = Rs.GetRows(1)
		Else
			Set rstmp = Execute("SELECT skinid,page_content,page_setting FROM [ECCMS_Template] WHERE ChannelID = " & ChannelID & " And isDefault = 1 And pageid = " & pageid)
			Value = rstmp.GetRows(1)
			Set rstmp = Nothing
		End If
		Set Rs = Nothing
	End Sub
	Private Sub TemplatesMainCache(skinid)
		On Error Resume Next
		Dim Rs, SQL, rstmp
		SQL = "SELECT TemplateDir,page_content,page_setting FROM [ECCMS_Template] WHERE pageid = 0 And skinid = " & skinid & " And ChannelID = 0"
		Set Rs = Execute(SQL)
		If Not Rs.EOF Then
			Value = Rs.GetRows(1)
		Else
			Set rstmp = Execute("SELECT TemplateDir,page_content,page_setting from [ECCMS_Template] WHERE pageid = 0 And isDefault = 1 And ChannelID = 0")
			Value = rstmp.GetRows(1)
			Set rstmp = Nothing
		End If
		Set Rs = Nothing
	End Sub
	
	Public Property Let ByValue(ByVal vNewValue)
		Dim tmpstr
		tmpstr = vNewValue
		Html_Setting = tmpstr(2, 0)
		Html_Setting = Split(Html_Setting, "|||")
		HtmlContent = tmpstr(1, 0)
		
		If CInt(Html_Setting(0)) <> 0 Then
			HtmlContent = HtmlTop & HtmlContent & HtmlFoot
		End If
		
		'用户信息
		dim userimg
		userimg=enchiasp.InstallDir & "user/images/icon/user_manager.gif"
		if enchiasp.membername="" then
			HtmlContent = Replace(HtmlContent,"{$showuserinfo}", Main_Setting(27))
		HtmlContent = Replace(HtmlContent,"{$showuserinfo2}", Main_Setting(33))
		else
			HtmlContent = Replace(HtmlContent ,"{$showuserinfo}","<table border='0' width='80%' align='center'><tr><td><img src='"& userimg &"'>欢迎您 <font color=red>"&enchiasp.membername&"</font>  "& enchiasp.membergroup &"</td></tr><tr><td>"&"<a href='"& enchiasp.InstallDir&"user/index.asp'>·进入用户管理中心</a></td></tr><tr><td><a href='"& enchiasp.InstallDir &"user/logout.asp'>·退出登陆</a>"&"</td></tr></table>")
			HtmlContent = Replace(HtmlContent ,"{$showuserinfo2}","<table border='0' width='80%' align='center'><tr><td><img src='"& userimg &"'>欢迎您 <font color=red>"&enchiasp.membername&"</font>  "& enchiasp.membergroup &"</td><td>"&"<a href='"& enchiasp.InstallDir&"user/index.asp'>·进入用户管理中心</a></td><td><a href='"& enchiasp.InstallDir &"user/logout.asp'>·退出登陆</a>"&"</td></tr></table>")
		end if
	
		
		HtmlContent = Replace(HtmlContent, "{$Style_CSS}", HtmlCss)
		HtmlContent = Replace(HtmlContent, "{$SkinPath}", SkinPath)
		HtmlContent = Replace(HtmlContent, "{$Width}", Main_Setting(0))
		
	
		'将图片变化加载到通用的部分设置
		HtmlContent = Replace(HtmlContent,"{$tupianhuan}", Main_Setting(24))
		HtmlContent = Replace(HtmlContent,"{$tupianhuanpic}", enchiasp.tupianhuanpic)
		HtmlContent = Replace(HtmlContent,"{$tupianhuanlink}", enchiasp.tupianhuanlink)
		HtmlContent = Replace(HtmlContent,"{$tupianhuantext}", enchiasp.tupianhuantext)


HtmlContent = Replace(HtmlContent,"{$tupianhuan2}", Main_Setting(30))
		HtmlContent = Replace(HtmlContent,"{$tupianhuanpic2}", enchiasp.tupianhuanpic2)
		HtmlContent = Replace(HtmlContent,"{$tupianhuanlink2}", enchiasp.tupianhuanlink2)
		HtmlContent = Replace(HtmlContent,"{$tupianhuantext2}", enchiasp.tupianhuantext2)


HtmlContent = Replace(HtmlContent,"{$tupianhuan3}", Main_Setting(31))
		HtmlContent = Replace(HtmlContent,"{$tupianhuanpic3}", enchiasp.tupianhuanpic3)
		HtmlContent = Replace(HtmlContent,"{$tupianhuanlink3}", enchiasp.tupianhuanlink3)
		HtmlContent = Replace(HtmlContent,"{$tupianhuantext3}", enchiasp.tupianhuantext3)


HtmlContent = Replace(HtmlContent,"{$tupianhuan4}", Main_Setting(32))
		HtmlContent = Replace(HtmlContent,"{$tupianhuanpic4}", enchiasp.tupianhuanpic4)
		HtmlContent = Replace(HtmlContent,"{$tupianhuanlink4}", enchiasp.tupianhuanlink4)
		HtmlContent = Replace(HtmlContent,"{$tupianhuantext4}", enchiasp.tupianhuantext4)


HtmlContent = Replace(HtmlContent,"{$tupianhuan5}", Main_Setting(34))
		HtmlContent = Replace(HtmlContent,"{$tupianhuanpic5}", enchiasp.tupianhuanpic5)
		HtmlContent = Replace(HtmlContent,"{$tupianhuanlink5}", enchiasp.tupianhuanlink5)
		HtmlContent = Replace(HtmlContent,"{$tupianhuantext5}", enchiasp.tupianhuantext5)


HtmlContent = Replace(HtmlContent,"{$tupianhuan6}", Main_Setting(35))
		HtmlContent = Replace(HtmlContent,"{$tupianhuanpic6}", enchiasp.tupianhuanpic6)
		HtmlContent = Replace(HtmlContent,"{$tupianhuanlink6}", enchiasp.tupianhuanlink6)
		HtmlContent = Replace(HtmlContent,"{$tupianhuantext6}", enchiasp.tupianhuantext6)






		'底部滚动图片
		HtmlContent = Replace(HtmlContent,"{$dibuhuan}", Main_Setting(25))
		HtmlContent = Replace(HtmlContent,"{$dibutupian}", enchiasp.dibutupian)
		
		
		'首页VOD点播
		HtmlContent = Replace(HtmlContent,"{$vod}", Main_Setting(26))
		HtmlContent = Replace(HtmlContent,"{$vodpath}", enchiasp.vodpath)
		
	
	

		


		HtmlContent = Replace(HtmlContent, "{$ChannelMenu}", ChannelMenu)
		HtmlContent = Replace(HtmlContent, "{$WebSiteName}", SiteName)
		HtmlContent = Replace(HtmlContent, "{$WebSiteUrl}", SiteUrl)
		HtmlContent = Replace(HtmlContent, "{$MasterMail}", MasterMail)
		HtmlContent = Replace(HtmlContent, "{$Keyword}", keywords)
		HtmlContent = Replace(HtmlContent, "{$Copyright}", Copyright)
		HtmlContent = Replace(HtmlContent, "{$IndexName}", IndexName)
		HtmlContent = Replace(HtmlContent, "{$Version}", "")
		HtmlContent = HtmlContent
	End Property
	Public Property Get ByValue()
		ByValue = HtmlContent
	End Property
	Public Property Let HTMLValue(ByVal vNewValue)
		Dim TempStr
		TempStr = vNewValue
		'用户信息
		dim userimg
		userimg=enchiasp.InstallDir & "user/images/icon/user_manager.gif"
		if enchiasp.membername="" then
			HtmlContent = Replace(HtmlContent,"{$showuserinfo}", Main_Setting(27))
		else
			HtmlContent = Replace(HtmlContent ,"{$showuserinfo}","<table border='0' width='80%' align='center'><tr><td><img src='"& userimg &"'>欢迎您 <font color=red>"&enchiasp.membername&"</font>  "& enchiasp.membergroup &"</td><tr><td>"&"<a href='"& enchiasp.InstallDir&"user/index.asp'>·进入用户管理中心</a></td></tr><tr><td><a href='"& enchiasp.InstallDir & "bbs/' target='_blank'>·进入论坛</a></td></tr><tr><td><a href='"& enchiasp.InstallDir &"user/logout.asp'>·退出登陆</a>"&"</td></tr></table>")
		end if
		TempStr = Replace(TempStr, "{$Style_CSS}", HtmlCss)
		TempStr = Replace(TempStr, "{$SkinPath}", SkinPath)
		TempStr = Replace(TempStr, "{$Width}", Main_Setting(0))
		'将图片变化加载到通用的部分设置
		HtmlContent = Replace(HtmlContent,"{$tupianhuan}", Main_Setting(24))
		HtmlContent = Replace(HtmlContent,"{$tupianhuanpic}", enchiasp.tupianhuanpic)
		HtmlContent = Replace(HtmlContent,"{$tupianhuanlink}", enchiasp.tupianhuanlink)
		HtmlContent = Replace(HtmlContent,"{$tupianhuantext}", enchiasp.tupianhuantext)
		



HtmlContent = Replace(HtmlContent,"{$tupianhuan2}", Main_Setting(30))
		HtmlContent = Replace(HtmlContent,"{$tupianhuanpic2}", enchiasp.tupianhuanpic2)
		HtmlContent = Replace(HtmlContent,"{$tupianhuanlink2}", enchiasp.tupianhuanlink2)
		HtmlContent = Replace(HtmlContent,"{$tupianhuantext2}", enchiasp.tupianhuantext2)


HtmlContent = Replace(HtmlContent,"{$tupianhuan3}", Main_Setting(31))
		HtmlContent = Replace(HtmlContent,"{$tupianhuanpic3}", enchiasp.tupianhuanpic3)
		HtmlContent = Replace(HtmlContent,"{$tupianhuanlink3}", enchiasp.tupianhuanlink3)
		HtmlContent = Replace(HtmlContent,"{$tupianhuantext3}", enchiasp.tupianhuantext3)


HtmlContent = Replace(HtmlContent,"{$tupianhuan4}", Main_Setting(32))
		HtmlContent = Replace(HtmlContent,"{$tupianhuanpic4}", enchiasp.tupianhuanpic4)
		HtmlContent = Replace(HtmlContent,"{$tupianhuanlink4}", enchiasp.tupianhuanlink4)
		HtmlContent = Replace(HtmlContent,"{$tupianhuantext4}", enchiasp.tupianhuantext4)


HtmlContent = Replace(HtmlContent,"{$tupianhuan5}", Main_Setting(34))
		HtmlContent = Replace(HtmlContent,"{$tupianhuanpic5}", enchiasp.tupianhuanpic5)
		HtmlContent = Replace(HtmlContent,"{$tupianhuanlink5}", enchiasp.tupianhuanlink5)
		HtmlContent = Replace(HtmlContent,"{$tupianhuantext5}", enchiasp.tupianhuantext5)


HtmlContent = Replace(HtmlContent,"{$tupianhuan6}", Main_Setting(35))
		HtmlContent = Replace(HtmlContent,"{$tupianhuanpic6}", enchiasp.tupianhuanpic6)
		HtmlContent = Replace(HtmlContent,"{$tupianhuanlink6}", enchiasp.tupianhuanlink6)
		HtmlContent = Replace(HtmlContent,"{$tupianhuantext6}", enchiasp.tupianhuantext6)









		'底部滚动图片
		HtmlContent = Replace(HtmlContent,"{$dibuhuan}", Main_Setting(25))
		HtmlContent = Replace(HtmlContent,"{$dibutupian}", enchiasp.dibutupian)
		
		'首页VOD点播
		HtmlContent = Replace(HtmlContent,"{$vod}", Main_Setting(26))
		HtmlContent = Replace(HtmlContent,"{$vodpath}", enchiasp.vodpath)



		
		TempStr = Replace(TempStr, "{$ChannelMenu}", ChannelMenu)
		TempStr = Replace(TempStr, "{$WebSiteName}", SiteName)
		TempStr = Replace(TempStr, "{$WebSiteUrl}", SiteUrl)
		TempStr = Replace(TempStr, "{$MasterMail}", MasterMail)
		TempStr = Replace(TempStr, "{$Keyword}", keywords)
		TempStr = Replace(TempStr, "{$Copyright}", Copyright)
		TempStr = Replace(TempStr, "{$IndexName}", IndexName)
		TempStr = Replace(TempStr, "{$Version}", "")
		sHtmlContent = TempStr
	End Property
	Public Property Get HTMLValue()
		HTMLValue = sHtmlContent
	End Property
	Public Property Get HtmlSetting(n)
		HtmlSetting = Html_Setting(n)
	End Property
	Public Property Get MainSetting(n)
		MainSetting = Main_Setting(n)
	End Property
	'================================================
	'过程名：GetSiteUrl
	'作  用：取得带端口的URL
	'================================================
	Public Property Get GetSiteUrl()
		If Request.ServerVariables("SERVER_PORT") = "80" Then
			GetSiteUrl = "http://" & Request.ServerVariables("server_name")
		Else
			GetSiteUrl = "http://" & Request.ServerVariables("server_name") & ":" & Request.ServerVariables("SERVER_PORT")
		End If
	End Property
	'================================================
	'函数名：FormEncode
	'作  用：过虑提交的表单数据
	'参  数：str ----原字符串  n ----字符长度
	'================================================
	Public Function FormEncode(ByVal str, ByVal n)
		If Not IsNull(str) And Trim(str) <> "" Then
			str = Left(str, n)
			str = Replace(str, ">", "&gt;")
			str = Replace(str, "<", "&lt;")
			str = Replace(str, "&#62;", "&gt;")
			str = Replace(str, "&#60;", "&lt;")
			str = Replace(str, "'", "&#39;")
			str = Replace(str, Chr(34), "&quot;")
			str = Replace(str, "%", "％")
			str = Replace(str, vbNewLine, "")
			FormEncode = Trim(str)
		Else
			FormEncode = ""
		End If
	End Function
	'================================================
	'函数名：ChkKeyWord
	'作  用：过滤关键字
	'参  数：keyword ----关键字
	'================================================
	Public Function ChkKeyWord(ByVal keyword)
		Dim FobWords, i
		On Error Resume Next
		FobWords = Array(91, 92, 304, 305, 430, 431, 437, 438, 12460, 12461, 12462, 12463, 12464, 12465, 12466, 12467, 12468, 12469, 12470, 12471, 12472, 12473, 12474, 12475, 12476, 12477, 12478, 12479, 12480, 12481, 12482, 12483, 12485, 12486, 12487, 12488, 12489, 12490, 12496, 12497, 12498, 12499, 12500, 12501, 12502, 12503, 12504, 12505, 12506, 12507, 12508, 12509, 12510, 12521, 12532, 12533, 65339, 65340)
		For i = 1 To UBound(FobWords, 1)
			If InStr(keyword, ChrW(FobWords(i))) > 0 Then
				keyword = Replace(keyword, ChrW(FobWords(i)), "")
			End If
		Next
		keyword = Left(keyword, 100)
		FobWords = Array("~", "!", "@", "#", "$", "%", "^", "&", "*", "(", ")", "_", "+", "=", "`", "[", "]", "{", "}", ";", ":", """", "'", ",", "<", ">", ".", "/", "\", "?", "_")
		For i = 0 To UBound(FobWords, 1)
			If InStr(keyword, FobWords(i)) > 0 Then
				keyword = Replace(keyword, FobWords(i), "")
			End If
		Next
		ChkKeyWord = keyword
	End Function
	'================================================
	'函数名：JAPEncode
	'作  用：日文片假名编码
	'参  数：str ----原字符
	'================================================
	Public Function JAPEncode(ByVal str)
		Dim FobWords, i
		On Error Resume Next
		If IsNull(str) Or Trim(str) = "" Then
			JAPEncode = ""
			Exit Function
		End If
		FobWords = Array(92, 304, 305, 430, 431, 437, 438, 12460, 12461, 12462, 12463, 12464, 12465, 12466, 12467, 12468, 12469, 12470, 12471, 12472, 12473, 12474, 12475, 12476, 12477, 12478, 12479, 12480, 12481, 12482, 12483, 12485, 12486, 12487, 12488, 12489, 12490, 12496, 12497, 12498, 12499, 12500, 12501, 12502, 12503, 12504, 12505, 12506, 12507, 12508, 12509, 12510, 12521, 12532, 12533, 65340)
		For i = 1 To UBound(FobWords, 1)
			If InStr(str, ChrW(FobWords(i))) > 0 Then
				str = Replace(str, ChrW(FobWords(i)), "&#" & FobWords(i) & ";")
			End If
		Next
		JAPEncode = str
	End Function
	'================================================
	'函数名：JAPUncode
	'作  用：日文片假名解码
	'参  数：str ----原字符
	'================================================
	Public Function JAPUncode(ByVal str)
		Dim FobWords, i
		On Error Resume Next
		If IsNull(str) Or Trim(str) = "" Then
			JAPUncode = ""
			Exit Function
		End If
		FobWords = Array(92, 304, 305, 430, 431, 437, 438, 12460, 12461, 12462, 12463, 12464, 12465, 12466, 12467, 12468, 12469, 12470, 12471, 12472, 12473, 12474, 12475, 12476, 12477, 12478, 12479, 12480, 12481, 12482, 12483, 12485, 12486, 12487, 12488, 12489, 12490, 12496, 12497, 12498, 12499, 12500, 12501, 12502, 12503, 12504, 12505, 12506, 12507, 12508, 12509, 12510, 12521, 12532, 12533, 65340)
		For i = 1 To UBound(FobWords, 1)
			If InStr(str, "&#" & FobWords(i) & ";") > 0 Then
				str = Replace(str, "&#" & FobWords(i) & ";", ChrW(FobWords(i)))
			End If
		Next
		str = Replace(str, Chr(0), "")
		str = Replace(str, "'", "''")
		JAPUncode = str
	End Function
	'=============================================================
	'函数作用：带脏话过滤
	'=============================================================
	Public Function ChkBadWords(ByVal str)
		If IsNull(str) Then Exit Function
		Dim i, Bwords, Bwordr
		Bwords = Split(Badwords, "|")
		Bwordr = Split(Badwordr, "|")
		For i = 0 To UBound(Bwords)
			If i > UBound(Bwordr) Then
				str = Replace(str, Bwords(i), "*")
			Else
				str = Replace(str, Bwords(i), Bwordr(i))
			End If
		Next
		ChkBadWords = str
	End Function
	'=============================================================
	'函数作用：过滤HTML代码，带脏话过滤
	'=============================================================
	Public Function HTMLEncode(ByVal fString)
		If Not IsNull(fString) Then
			fString = Replace(fString, ">", "&gt;")
			fString = Replace(fString, "<", "&lt;")
			fString = Replace(fString, Chr(32), " ")
			fString = Replace(fString, Chr(9), " ")
			fString = Replace(fString, Chr(34), "&quot;")
			fString = Replace(fString, Chr(39), "&#39;")
			fString = Replace(fString, Chr(13), "")
			fString = Replace(fString, " ", "&nbsp;")
			fString = Replace(fString, Chr(10), "<br /> ")
			fString = ChkBadWords(fString)
			HTMLEncode = fString
		End If
	End Function
	'=============================================================
	'函数作用：过滤HTML代码，不带脏话过滤
	'=============================================================
	Public Function HTMLEncodes(ByVal fString)
		If Not IsNull(fString) Then
			fString = Replace(fString, "'", "&#39;")
			fString = Replace(fString, ">", "&gt;")
			fString = Replace(fString, "<", "&lt;")
			fString = Replace(fString, Chr(32), " ")
			fString = Replace(fString, Chr(9), " ")
			fString = Replace(fString, Chr(34), "&quot;")
			fString = Replace(fString, Chr(39), "&#39;")
			fString = Replace(fString, Chr(13), "")
			fString = Replace(fString, Chr(10), "<br /> ")
			fString = Replace(fString, " ", "&nbsp;")
			HTMLEncodes = fString
		End If
	End Function
	'=============================================================
	'函数作用：判断发言是否来自外部
	'=============================================================
	Public Function CheckPost()
		On Error Resume Next
		Dim server_v1, server_v2
		CheckPost = False
		server_v1 = CStr(Request.ServerVariables("HTTP_REFERER"))
		server_v2 = CStr(Request.ServerVariables("SERVER_NAME"))
		If Mid(server_v1, 8, Len(server_v2)) = server_v2 Then
			CheckPost = True
		End If
	End Function
	'=============================================================
	'函数作用：判断来源URL是否来自外部
	'=============================================================
	Public Function CheckOuterUrl()
		On Error Resume Next
		Dim server_v1, server_v2
		server_v1 = Replace(LCase(Trim(Request.ServerVariables("HTTP_REFERER"))), "http://", "")
		server_v2 = LCase(Trim(Request.ServerVariables("SERVER_NAME")))


		If  server_v1 <> "" And Left(server_v1, Len(server_v2)) <> server_v2 Then
			CheckOuterUrl = False
		Else
			CheckOuterUrl = True
		End If
		'If  InStr(server_v1, "ltzxw.com") <= 0 and server_v1 <> "" And Left(server_v1, Len(server_v2)) <> server_v2 Then
			'CheckOuterUrl = False
		'Else
			'CheckOuterUrl = True
		'End If
	End Function
	
	
	
	'================================================
	'函数名：GotTopic
	'作  用：显示字符串长度
	'参  数：str   ----原字符串
	'        strlen  ----显示字符长度
	'================================================
	Public Function GotTopic(ByVal str, ByVal strLen)
		Dim l, t, c, i
		Dim strTemp
		On Error Resume Next
		str = Trim(str)
		str = Replace(str, "&nbsp;", " ")
		str = Replace(str, "&gt;", ">")
		str = Replace(str, "&lt;", "<")
		str = Replace(str, "&#62;", ">")
		str = Replace(str, "&#60;", "<")
		str = Replace(str, "&#39;", "'")
		str = Replace(str, "&quot;", Chr(34))
		str = Replace(str, vbNewLine, "")
		l = Len(str)
		t = 0
		For i = 1 To l
			c = Abs(Asc(Mid(str, i, 1)))
			If c > 255 Then
				t = t + 2
			Else
				t = t + 1
			End If
			If t >= strLen Then
				strTemp = Left(str, i) & "..."
				Exit For
			Else
				strTemp = str & " "
			End If
		Next
		GotTopic = CheckTopic(strTemp)
	End Function
	Public Function CheckTopic(ByVal strContent)
		Dim re
		On Error Resume Next
		Set re = New RegExp
		re.IgnoreCase = True
		re.Global = True
		re.Pattern = "(<s+cript(.+?)<\/s+cript>)"
		strContent = re.Replace(strContent, "")
		re.Pattern = "(<iframe(.+?)<\/iframe>)"
		strContent = re.Replace(strContent, "")
		re.Pattern = "(&#62;)"
		strContent = re.Replace(strContent, "&gt;")
		re.Pattern = "(&#60;)"
		strContent = re.Replace(strContent, "&lt;")
		Set re = Nothing
		strContent = Replace(strContent, ">", "&gt;")
		strContent = Replace(strContent, "<", "&lt;")
		strContent = Replace(strContent, "'", "&#39;")
		strContent = Replace(strContent, Chr(34), "&quot;")
		strContent = Replace(strContent, "%", "％")
		strContent = Replace(strContent, vbNewLine, "")
		CheckTopic = Trim(strContent)
	End Function
	'================================================
	'函数名：ReadTopic
	'作  用：显示字符串长度
	'参  数：str   ----原字符串
	'        strlen  ----显示字符长度
	'================================================
	Public Function ReadTopic(ByVal str, ByVal strLen)
		Dim l, t, c, i
		On Error Resume Next
		str = Replace(str, "&nbsp;", " ")
		If Len(str) < strLen Then
			str = str & String(strLen - Len(str), ".")
		Else
			str = str
		End If
		l = Len(str)
		t = 0
		For i = 1 To l
			c = Abs(Asc(Mid(str, i, 1)))
			If c > 255 Then
				t = t + 2
			Else
				t = t + 1
			End If
			If t >= strLen Then
				ReadTopic = Left(str, i) & "..."
				Exit For
			Else
				ReadTopic = str & "..."
			End If
		Next
	End Function
	'================================================
	'函数名：strLength
	'作  用：计字符串长度
	'参  数：str   ----字符串
	'================================================
	Public Function strLength(ByVal str)
		On Error Resume Next
		If IsNull(str) Or str = "" Then
			strLength = 0
			Exit Function
		End If
		Dim WINNT_CHINESE
		WINNT_CHINESE = (Len("例子") = 2)
		If WINNT_CHINESE Then
			Dim l, t
			Dim i, c
			l = Len(str)
			t = l
			For i = 1 To l
				c = Asc(Mid(str, i, 1))
				If c < 0 Then c = c + 65536
				If c > 255 Then t = t + 1
			Next
			strLength = t
		Else
			strLength = Len(str)
		End If
	End Function
	'=================================================
	'函数名：isInteger
	'作  用：判断数字是否整型
	'参  数：para ----参数
	'=================================================
	Public Function isInteger(ByVal para)
		On Error Resume Next
		Dim str
		Dim l, i
		If IsNull(para) Then
			isInteger = False
			Exit Function
		End If
		str = CStr(para)
		If Trim(str) = "" Then
			isInteger = False
			Exit Function
		End If
		l = Len(str)
		For i = 1 To l
			If Mid(str, i, 1) > "9" Or Mid(str, i, 1) < "0" Then
				isInteger = False
				Exit Function
			End If
		Next
		isInteger = True
		If Err.Number <> 0 Then Err.Clear
	End Function
	Public Function CutString(ByVal str, ByVal strLen)
		On Error Resume Next
		
		Dim HtmlStr, l, re, strContent
		
		HtmlStr = str
		HtmlStr = Replace(HtmlStr, "&nbsp;", " ")
		HtmlStr = Replace(HtmlStr, "&quot;", Chr(34))
		HtmlStr = Replace(HtmlStr, "&#39;", Chr(39))
		HtmlStr = Replace(HtmlStr, "&#123;", Chr(123))
		HtmlStr = Replace(HtmlStr, "&#125;", Chr(125))
		HtmlStr = Replace(HtmlStr, "&#36;", Chr(36))
		HtmlStr = Replace(HtmlStr, vbCrLf, "")
		HtmlStr = Replace(HtmlStr, "====", "")
		HtmlStr = Replace(HtmlStr, "----", "")
		HtmlStr = Replace(HtmlStr, "////", "")
		HtmlStr = Replace(HtmlStr, "\\\\", "")
		HtmlStr = Replace(HtmlStr, "####", "")
		HtmlStr = Replace(HtmlStr, "@@@@", "")
		HtmlStr = Replace(HtmlStr, "****", "")
		HtmlStr = Replace(HtmlStr, "~~~~", "")
		Set re = New RegExp
		re.IgnoreCase = True
		re.Global = True
		re.Pattern = "\[br\]"
		HtmlStr = re.Replace(HtmlStr, "")
		re.Pattern = "\[align=right\](.*)\[\/align\]"
		HtmlStr = re.Replace(HtmlStr, "")
		re.Pattern = "<(.[^>]*)>"
		HtmlStr = re.Replace(HtmlStr, "")
		Set re = Nothing
		HtmlStr = Replace(HtmlStr, "&gt;", ">")
		HtmlStr = Replace(HtmlStr, "&lt;", "<")
		l = Len(HtmlStr)
		If l >= strLen Then
			strContent = Left(HtmlStr, strLen) & "..."
		Else
			strContent = HtmlStr & " "
		End If
		strContent = Replace(strContent, Chr(34), "&quot;")
		strContent = Replace(strContent, Chr(39), "&#39;")
		strContent = Replace(strContent, Chr(36), "&#36;")
		strContent = Replace(strContent, Chr(123), "&#123;")
		strContent = Replace(strContent, Chr(125), "&#125;")
		strContent = Replace(strContent, ">", "&gt;")
		strContent = Replace(strContent, "<", "&lt;")
		CutString = strContent
	End Function
	'================================================
	'函数名：CheckInfuse
	'作  用：防止SQL注入
	'参  数：str   ----原字符串
	'        strLen  ----提交字符串长度
	'================================================
	Public Function CheckInfuse(ByVal str, ByVal strLen)
		Dim strUnsafe, arrUnsafe
		Dim i
		
		If Trim(str) = "" Then
			CheckInfuse = ""
			Exit Function
		End If
		str = Left(str, strLen)
		
		On Error Resume Next
		strUnsafe = "'|^|;|and|exec|insert|select|delete|update|count|*|%|chr|mid|master|truncate|char|declare"
		If Trim(str) <> "" Then
			If Len(str) > strLen Then
				Response.Write "<Script Language=JavaScript>alert('安全系统提示↓\n\n您提交的字符数超过了限制！');history.back(-1)</Script>"
				CheckInfuse = ""
				Response.End
			End If
			arrUnsafe = Split(strUnsafe, "|")
			For i = 0 To UBound(arrUnsafe)
				If InStr(1, str, arrUnsafe(i), 1) > 0 Then
					Response.Write "<Script Language=JavaScript>alert('安全系统提示↓\n\n请不要在参数中包含非法字符！');history.back(-1)</Script>"
					CheckInfuse = ""
					Response.End
				End If
			Next
		End If
		CheckInfuse = Trim(str)
		Exit Function
		If Err.Number <> 0 Then
			Err.Clear
			Response.Write "<Script Language=JavaScript>alert('安全系统提示↓\n\n请不要在参数中包含非法字符！');history.back(-1)</Script>"
			CheckInfuse = ""
			Response.End
		End If
	End Function
	Public Sub PreventInfuse()
		On Error Resume Next
		Dim SQL_Nonlicet, arrNonlicet
		Dim PostRefer, GetRefer, Sql_DATA
		
		SQL_Nonlicet = "'|;|^|and|exec|insert|select|delete|update|count|*|%|chr|mid|master|truncate|char|declare"
		arrNonlicet = Split(SQL_Nonlicet, "|")
		If Request.Form <> "" Then
			For Each PostRefer In Request.Form
				For Sql_DATA = 0 To UBound(arrNonlicet)
					If InStr(1, Request.Form(PostRefer), arrNonlicet(Sql_DATA), 1) > 0 Then
					Response.Write "<Script Language=JavaScript>alert('安全系统提示↓\n\n请不要在参数中包含非法字符！');history.back(-1)</Script>"
					Response.End
					End If
				Next
			Next
		End If

		If Request.QueryString <> "" Then
			For Each GetRefer In Request.QueryString
				For Sql_DATA = 0 To UBound(arrNonlicet)
					If InStr(1, Request.QueryString(GetRefer), arrNonlicet(Sql_DATA), 1) > 0 Then
					Response.Write "<Script Language=JavaScript>alert('安全系统提示↓\n\n请不要在参数中包含非法字符！');history.back(-1)</Script>"
					Response.End
					End If
				Next
			Next
		End If
	End Sub
	'================================================
	'函数名：ChkQueryStr
	'作  用：过虑查询的非法字符
	'参  数：str   ----原字符串
	'返回值：过滤后的字符
	'================================================
	Public Function ChkQueryStr(ByVal str)
		On Error Resume Next
		If IsNull(str) Then
			ChkQueryStr = ""
			Exit Function
		End If
		str = Replace(str, "!", "")
		str = Replace(str, "]", "")
		str = Replace(str, "[", "")
		str = Replace(str, ")", "")
		str = Replace(str, "(", "")
		str = Replace(str, "|", "")
		str = Replace(str, "+", "")
		str = Replace(str, "=", "")
		str = Replace(str, "'", "''")
		str = Replace(str, "%", "")
		str = Replace(str, "&", "")
		str = Replace(str, "#", "")
		str = Replace(str, "^", "")
		str = Replace(str, "&nbsp;", " ")
		str = Replace(str, Chr(37), "")
		str = Replace(str, Chr(0), "")
		ChkQueryStr = str
	End Function
	'================================================
	'过程名：CheckQuery
	'作  用：限制搜索的关键字
	'参  数：str ----搜索的字符串
	'返回值：True; False
	'================================================
	Public Function CheckQuery(ByVal str)
		Dim FobWords, i, keyword
		keyword = str
		On Error Resume Next
		FobWords = Array(91, 92, 304, 305, 430, 431, 437, 438, 12460, 12461, 12462, 12463, 12464, 12465, 12466, 12467, 12468, 12469, 12470, 12471, 12472, 12473, 12474, 12475, 12476, 12477, 12478, 12479, 12480, 12481, 12482, 12483, 12485, 12486, 12487, 12488, 12489, 12490, 12496, 12497, 12498, 12499, 12500, 12501, 12502, 12503, 12504, 12505, 12506, 12507, 12508, 12509, 12510, 12532, 12533, 65339, 65340)
		For i = 1 To UBound(FobWords, 1)
			If InStr(keyword, ChrW(FobWords(i))) > 0 Then
				CheckQuery = False
				Exit Function
			End If
		Next
		FobWords = Array("~", "!", "@", "#", "$", "%", "^", "&", "*", "(", ")", "_", "+", "=", "`", "[", "]", "{", "}", ";", ":", """", "'", "<", ">", ".", "/", "\", "|", "?", "about", "after", "all", "also", "an", "and", "another", "any", "are", "as", "at", "be", "because", "been", "before", "being", "between", "both", "but", "by", "came", "can", "come", "could", "did", "do", "each", "for", "from", "get", "got", "had", "has", "have", "he", "her", "here", "him", "himself", "his", "how", "if", "in", "into", "is", "it", "like", "make", "many", "me", "might", "more", "most", "much", "must", "my", "never", "now", "of", "on", "only", "or", "other", "our", "out", "over", "said", "same", "see", "should", "since", "some", "still", "such", "take", "than", "that", "the", "their", "them", "then", "there", "these", "they", "this")
		keyword = Left(keyword, 100)
		keyword = Replace(keyword, "!", " ")
		keyword = Replace(keyword, "]", " ")
		keyword = Replace(keyword, "[", " ")
		keyword = Replace(keyword, ")", " ")
		keyword = Replace(keyword, "(", " ")
		keyword = Replace(keyword, "　", " ")
		keyword = Replace(keyword, "-", " ")
		keyword = Replace(keyword, "/", " ")
		keyword = Replace(keyword, "+", " ")
		keyword = Replace(keyword, "=", " ")
		keyword = Replace(keyword, ",", " ")
		keyword = Replace(keyword, "'", " ")
		For i = 0 To UBound(FobWords, 1)
			If keyword = FobWords(i) Then
				CheckQuery = False
				Exit Function
			End If
		Next
		CheckQuery = True
	End Function
	'================================================
	'函数名：IsValidStr
	'作  用：判断字符串中是否含有非法字符
	'参  数：str   ----原字符串
	'返回值：False,True -----布尔值
	'================================================
	Public Function IsValidStr(ByVal str)
		IsValidStr = False
		On Error Resume Next
		If IsNull(str) Then Exit Function
		If Trim(str) = Empty Then Exit Function
		Dim ForbidStr, i
		ForbidStr = "and|chr|:|=|%|&|$|#|@|+|-|*|/|\|<|>|;|,|^|" & Chr(32) & "|" & Chr(34) & "|" & Chr(39) & "|" & Chr(9)
		ForbidStr = Split(ForbidStr, "|")
		For i = 0 To UBound(ForbidStr)
			If InStr(1,str, ForbidStr(i),1) > 0 Then
				IsValidStr = False
				Exit Function
			End If
		Next
		IsValidStr = True
	End Function
	'================================================
	'函数名：IsValidPassword
	'作  用：判断密码中是否含有非法字符
	'参  数：str   ----原字符串
	'返回值：False,True -----布尔值
	'================================================
	Public Function IsValidPassword(ByVal str)
		IsValidPassword = False
		On Error Resume Next
		If IsNull(str) Then Exit Function
		If Trim(str) = Empty Then Exit Function
		Dim ForbidStr, i
		ForbidStr = "=and|chr|*|^|%|&|;|,|" & Chr(32) & "|" & Chr(34) & "|" & Chr(39) & "|" & Chr(9)
		ForbidStr = Split(ForbidStr, "|")
		For i = 0 To UBound(ForbidStr)
			If InStr(1, str, ForbidStr(i), 1) > 0 Then
				IsValidPassword = False
				Exit Function
			End If
		Next
		IsValidPassword = True
	End Function
	'================================================
	'函数名：IsValidChar
	'作  用：判断字符串中是否含有非法字符和中文
	'参  数：str   ----原字符串
	'返回值：False,True -----布尔值
	'================================================
	Public Function IsValidChar(ByVal str)
		IsValidChar = False
		On Error Resume Next
		
		If IsNull(str) Then Exit Function
		If Trim(str) = Empty Then Exit Function
		Dim ValidStr
		Dim i, l, s, c
		
		ValidStr = "ABCDEFGHIJKLMNOPQRSTUVWXYZ.-_:~\/0123456789"
		l = Len(str)
		s = UCase(str)
		For i = 1 To l
			c = Mid(s, i, 1)
			If InStr(ValidStr, c) = 0 Then
				IsValidChar = False
				Exit Function
			End If
		Next
		IsValidChar = True
	End Function
	'================================================
	'函数名：FormatDate
	'作  用：格式化日期
	'参  数：DateAndTime   ----原日期和时间
	'        para   ----日期格式
	'返回值：格式化后的日期
	'================================================
	Public Function FormatDate(DateAndTime, para)
		On Error Resume Next
		Dim y, m, d, h, mi, s, strDateTime
		FormatDate = DateAndTime
		If Not IsNumeric(para) Then Exit Function
		If Not IsDate(DateAndTime) Then Exit Function
		y = CStr(Year(DateAndTime))
		m = CStr(Month(DateAndTime))
		If Len(m) = 1 Then m = "0" & m
		d = CStr(Day(DateAndTime))
		If Len(d) = 1 Then d = "0" & d
		h = CStr(Hour(DateAndTime))
		If Len(h) = 1 Then h = "0" & h
		mi = CStr(Minute(DateAndTime))
		If Len(mi) = 1 Then mi = "0" & mi
		s = CStr(Second(DateAndTime))
		If Len(s) = 1 Then s = "0" & s
		Select Case para
		Case "1"
			strDateTime = y & "-" & m & "-" & d & " " & h & ":" & mi & ":" & s
		Case "2"
			strDateTime = y & "-" & m & "-" & d
		Case "3"
			strDateTime = y & "/" & m & "/" & d
		Case "4"
			strDateTime = y & "年" & m & "月" & d & "日"
		Case "5"
			strDateTime = m & "-" & d
		Case "6"
			strDateTime = m & "/" & d
		Case "7"
			strDateTime = m & "月" & d & "日"
		Case "8"
			strDateTime = y & "年" & m & "月"
		Case "9"
			strDateTime = y & "-" & m
		Case "10"
			strDateTime = y & "/" & m
		Case Else
			strDateTime = DateAndTime
		End Select
		FormatDate = strDateTime
	End Function
	'================================================
	'函数名：ReadFontMode
	'作  用：读取字体模式
	'参  数：str   ----原字符串
	'        vColor   -----颜色的值
	'        vFont   -----字体的值
	'返回值：新字符串
	'================================================
	Public Function ReadFontMode(str, vColor, vFont)
		Dim FontStr, tColor
		Dim ColorStr, arrColor
		
		If IsNull(str) Then
			ReadFontMode = ""
			Exit Function
		End If
		ReadFontMode = str
		On Error Resume Next
		If Not IsNumeric(vColor) Then Exit Function
		If Not IsNumeric(vFont) Then Exit Function
		
		Select Case CInt(vFont)
			Case 1
				FontStr = "<b>" & str & "</b>"
			Case 2
				FontStr = "<em>" & str & "</em>"
			Case 3
				FontStr = "<u>" & str & "</u>"
			Case 4
				FontStr = "<b><em>" & str & "</em></b>"
			Case 5
				FontStr = "<b><u>" & str & "</u></b>"
			Case 6
				FontStr = "<em><u>" & str & "</u></em>"
			Case 7
				FontStr = "<b><em><u>" & str & "</u></em></b>"
		Case Else
			FontStr = str
		End Select
		ReadFontMode = FontStr
		
		If vColor = "" Or vColor = 0 Then Exit Function
		ColorStr = "," & InitTitleColor
		arrColor = Split(ColorStr, ",")
		If vColor > UBound(arrColor) Then Exit Function
		tColor = Trim(arrColor(vColor))
		ReadFontMode = "<font color=" & tColor & ">" & FontStr & "</font>"
	End Function
	'=============================================================
	'函数名：ShowDateTime
	'作  用：读取日期格式
	'参  数：DateAndTime ---- 当前时间
	'        para ---- 时间格式
	'=============================================================
	Public Function ShowDateTime(DateAndTime, para)
		ShowDateTime = ""
		Dim strDate
		If Not IsDate(DateAndTime) Then Exit Function
		If DateAndTime >= Date Then
			strDate = "<font color='" & Main_Setting(1) & "'>"
			strDate = strDate & FormatDate(DateAndTime, para)
			strDate = strDate & "</font>"
		Else
			strDate = "<font color='" & Main_Setting(2) & "'>"
			strDate = strDate & FormatDate(DateAndTime, para)
			strDate = strDate & "</font>"
		End If
		ShowDateTime = strDate
	End Function
	Public Function ShowDatePath(strval, n)
		ShowDatePath = ""
		If Trim(strval) = "" Then Exit Function
		Dim strTempPath, strTime
		Dim y, m, d
		
		strTime = Left(strval, 8)
		y = Left(strTime, 4)
		m = Mid(strTime, 5, 2)
		d = Right(strTime, 2)
		Select Case CInt(n)
			Case 1
				strTempPath = y & "/" & m & "/" & d & "/"
			Case 2
				strTempPath = y & "/" & m & "/"
			Case 3
				strTempPath = y & m & "/"
			Case 4
				strTempPath = y & "/"
			Case 5
				strTempPath = y & "-" & m & "-" & d & "/"
			Case 6
				strTempPath = y & "-" & m & "/"
			Case 7
				strTempPath = "html/"
			Case 8
				strTempPath = "show/"
		Case Else
			strTempPath = ""
		End Select
		strTempPath = Replace(strTempPath, " ", "")
		ShowDatePath = CStr(strTempPath)
	End Function
	'=============================================================
	'函数名：ReadBriefTopicffd
	'作  用：读取简短标题
	'参  数：para
	'返回值：简短标题
	'=============================================================
	Public Function ReadBriefTopic(ByVal para)
		Dim sBriefTopic
		
		ReadBriefTopic = ""
		If Not IsNumeric(para) Then Exit Function
		If para = 0 Then Exit Function
		Select Case para
		Case "1"
			sBriefTopic = "<font color='blue'>[图文]</font>"
		Case "2"
			sBriefTopic = "<font color='red'>[组图]</font>"
		Case "3"
			sBriefTopic = "<font color='green'>[新闻]</font>"
		Case "4"
			sBriefTopic = "<font color='blue'>[推荐]</font>"
		Case "5"
			sBriefTopic = "<font color='red'>[注意]</font>"
		Case "6"
			sBriefTopic = "<font color='green'>[转载]</font>"
		Case Else
			sBriefTopic = ""
		End Select
		ReadBriefTopic = sBriefTopic
	End Function
	'=============================================================
	'函数名：ReadPicTopic
	'作  用：读取简短标题
	'参  数：para
	'返回值：简短标题
	'=============================================================
	Public Function ReadPicTopic(ByVal para)
		Dim sBriefTopic
		ReadPicTopic = ""
		If Not IsNumeric(para) Then Exit Function
		If para = 0 Then Exit Function
		Select Case para
		Case "1"
			sBriefTopic = "<font color='" & Main_Setting(4) & "'>[图文]</font>"
		Case "2"
			sBriefTopic = "<font color='" & Main_Setting(5) & "'>[组图]</font>"
		Case "3"
			sBriefTopic = "<font color='" & Main_Setting(6) & "'>[新闻]</font>"
		Case "4"
			sBriefTopic = "<font color='" & Main_Setting(4) & "'>[推荐]</font>"
		Case "5"
			sBriefTopic = "<font color='" & Main_Setting(5) & "'>[注意]</font>"
		Case "6"
			sBriefTopic = "<font color='" & Main_Setting(6) & "'>[转载]</font>"
		Case Else
			sBriefTopic = ""
		End Select
		ReadPicTopic = sBriefTopic
	End Function
	'=============================================================
	'函数名：ReadPayMoney
	'作  用：读取要支付的金钱
	'参  数：money   ----实际金钱
	'返回值：加上手续费后的金钱
	'=============================================================
	Public Function ReadPayMoney(ByVal money, ByVal Reduce)
		On Error Resume Next
		If money = 0 Then
			ReadPayMoney = 0
			Exit Function
		End If
		Dim arrChinaeBank, valPercent, Percents
		
		arrChinaeBank = Split(ChinaeBank, "|||")
		Percents = CCur(arrChinaeBank(2) / 100)
		
		If Percents = 0 Then
			ReadPayMoney = CCur(money)
		Else
			If CBool(Reduce) = True Then
				valPercent = Round(CCur(money) / (1 + 1 * Percents), 2)
				ReadPayMoney = CCur(valPercent)
			Else
				valPercent = Round(CCur(money) * Percents, 2)
				ReadPayMoney = CCur(money + valPercent)
			End If
		End If
	End Function
	'=============================================================
	'函数名：RebateMoney
	'作  用：读取打折的后金钱
	'参  数：money   ----实际金钱
	'        Discount   ----折扣
	'=============================================================
	Public Function RebateMoney(ByVal money, ByVal Discount)
		On Error Resume Next
		Dim Rebate
		
		money = CheckNumeric(money)
		Discount = CheckNumeric(Discount)
		If Discount > 0 And Discount < 10 Then
			Rebate = Round(money * (Discount / 10), 2)
			RebateMoney = CCur(Rebate)
		Else
			RebateMoney = CCur(money)
		End If
	End Function
	'================================================
	'函数名：Supplemental
	'作  用：补足参数
	'参  数：para ----原参数
	'        n ----增补的位数
	'================================================
	Public Function Supplemental(para, n)
		Supplemental = ""
		If Not IsNumeric(para) Then Exit Function
		If Len(para) < n Then
			Supplemental = String(n - Len(para), "0") & para
		Else
			Supplemental = para
		End If
	End Function
	'-----------------------------------------------------------------
	Public Function GetChannelDir(ByVal chanid)
		On Error Resume Next
		If Not IsNumeric(chanid) Then chanid = 1
		Name = "Channel" & chanid
		If ObjIsEmpty() Then ReloadChannel (chanid)
		CacheChannel = Value
		GetChannelDir = InstallDir & CacheChannel(2,0)
	End Function
	
	'================================================
	'函数名：GetImageUrl
	'作  用：获取图片URL
	'================================================
	Public Function GetImageUrl(ByVal url, ByVal ChannelDir)
		On Error Resume Next
		Dim strTempUrl, strImageUrl
		
		If Not IsNull(url) And Trim(url) <> "" And LCase(url) <> "http://" Then
			strTempUrl = InstallDir & ChannelDir
			If CheckUrl(url) = 1 Then
				strImageUrl = Trim(url)
			ElseIf CheckUrl(url) = 2 Then
				strImageUrl = url
			Else
				strImageUrl = Replace(url, "../", "")
				strImageUrl = Trim(strTempUrl & strImageUrl)
			End If
		Else
			strImageUrl = InstallDir & "images/no_pic.gif"
		End If
		GetImageUrl = strImageUrl
	End Function
	'-----------------------------------------------------------------
	'================================================
	'作  用：读取图片或者FLASH
	'参  数：url ----文件URL
	'        height ----高度
	'        width ----宽度
	'================================================
	Function GetFlashAndPic(ByVal url, ByVal height, ByVal width)
		On Error Resume Next
		Dim sExtName, ExtName, strTemp
		Dim strHeight, strWidth
		
		If Not IsNumeric(height) Or height < 1 Then
			strHeight = ""
		Else
			strHeight = " height=" & height
		End If
		If Not IsNumeric(width) Or width < 1 Then
			strWidth = ""
		Else
			strWidth = " width=" & width
		End If
		sExtName = Split(url, ".")
		ExtName = sExtName(UBound(sExtName))
		If LCase(ExtName) = "swf" Then
			strTemp = "<embed src=""" & url & """" & strWidth & strHeight & ">"
		Else
			strTemp = "<img src=""" & url & """" & strWidth & strHeight & " border=0>"
		End If
		GetFlashAndPic = strTemp
	End Function
	'================================================
	'函数名：ReadFileUrl
	'作  用：读取文件URL
	'================================================
	Public Function ReadFileUrl(url)
		On Error Resume Next
		ReadFileUrl = ""
		If url = "" Then Exit Function
		Dim strTemp
		If CheckUrl(url) = 1 Then
			strTemp = Trim(url)
		ElseIf CheckUrl(url) = 2 Then
			strTemp = Trim(url)
		Else
			strTemp = Replace(url, "../", "")
			strTemp = Trim(InstallDir & strTemp)
		End If
		ReadFileUrl = strTemp
	End Function
	Public Function CheckUrl(ByVal url)
		Dim strUrl
		If Left(url, 1) = "/" Then
			CheckUrl = 1
			Exit Function
		End If
		strUrl = LCase(Left(url, 6))
		Select Case Trim(strUrl)
		Case "http:/", "https:", "ftp://", "rtsp:/", "mms://"
			CheckUrl = 2
			Exit Function
		Case Else
			CheckUrl = 0
		End Select
	End Function
	'================================================
	'函数名：ReadFileName
	'作  用：读取HTML文件名
	'参  数：strname ----文件名称
	'        id ----数据ID
	'        ExtName ----HTML扩展名
	'        PrefixStr ----HTML名称前缀
	'        HtmlForm ----HTML文件格式
	'        n ----HTML分页
	'================================================
	Public Function ReadFileName(ByVal strname, ByVal id, ByVal ExtName, ByVal PrefixStr, ByVal HtmlForm, ByVal n)
		
		Dim strFileName, strExtName, CurrentPage
		If Trim(strname) = "" Then Exit Function
		If Trim(ExtName) = "" Then ExtName = ".html"
		If Not IsNumeric(n) Then n = 0
		On Error Resume Next
		If CInt(n) <= 1 Then
			CurrentPage = ""
		Else
			CurrentPage = "_" & n
		End If
		If Left(ExtName, 1) <> "." Then
			strExtName = "." & Trim(ExtName)
		Else
			strExtName = Trim(ExtName)
		End If
		Select Case Trim(HtmlForm)
			Case "1"
				strFileName = Trim(id)
			Case "2"
				strFileName = Trim(PrefixStr) & Trim(Supplemental(id, 3))
			Case "3"
				strFileName = Left(strname, 8)
				strFileName = strFileName & Trim(Supplemental(id, 3))
			Case "4"
				strFileName = Right(strname, 7)
				strFileName = strFileName & Trim(Supplemental(id, 3))
			Case Else
				strFileName = strname
		End Select
		strFileName = Replace(strFileName & CurrentPage & strExtName, " ", "")
		ReadFileName = CStr(strFileName)
	End Function
	'================================================
	'过程名：HtmlRndFileName
	'作  用：取HTML的随机文件名
	'================================================
	Function HtmlRndFileName()
		Dim sRnd
		Randomize
		sRnd = Int(90 * Rnd) + 10
		HtmlRndFileName = Replace(Replace(Replace(FormatDate(Now(), 1), "-", ""), ":", ""), " ", "") & sRnd
	End Function
	'================================================
	'函数名：ClassFileName
	'作  用：读取HTML文件列表名
	'参  数：ClassID ----分类ID
	'================================================
	Public Function ClassFileName(ByVal ClassID, ByVal ExtName, ByVal PrefixStr, ByVal n)
		Dim strFileName, strExtName, strClassID
		
		If Trim(ExtName) = "" Then ExtName = ".html"
		If Not IsNumeric(n) Then n = 0
		If Left(ExtName, 1) <> "." Then
			strExtName = "." & Trim(ExtName)
		Else
			strExtName = Trim(ExtName)
		End If
		If CInt(n) <= 1 Then
			strFileName = "index" & strExtName
		Else
			strClassID = Supplemental(ClassID, 3)
			strFileName = PrefixStr & strClassID & "_" & n & strExtName
		End If
		strFileName = Replace(strFileName, " ", "")
		ClassFileName = CStr(strFileName)
	End Function
	'================================================
	'函数名：SpecialFileName
	'作  用：读取专题HTML文件名
	'参  数：specid ----专题ID
	'================================================
	Public Function SpecialFileName(ByVal specid, ByVal ExtName, ByVal n)
		Dim strFileName, strExtName, strSpecialID
		
		If Trim(ExtName) = "" Then ExtName = ".html"
		If Not IsNumeric(n) Then n = 0
		If Left(ExtName, 1) <> "." Then
			strExtName = "." & Trim(ExtName)
		Else
			strExtName = Trim(ExtName)
		End If
		If CInt(n) <= 1 Then
			strFileName = "index" & strExtName
		Else
			strSpecialID = Supplemental(specid, 3)
			strFileName = "Special" & strSpecialID & "_" & n & strExtName
		End If
		strFileName = Replace(strFileName, " ", "")
		SpecialFileName = CStr(strFileName)
	End Function
	'================================================
	'函数名：ChannelMenu
	'作  用：显示频道菜单
	'================================================
	Public Function ChannelMenu()
		Dim SQL, Rs, i, TotalNumber,strTop
		Dim strContent, LinkTarget, ChannelName
		Dim ChannelUrl, HtmlContent, sCaption
		
		
		Name = "ChannelMenu"
		If ObjIsEmpty() Then
			If ChkNumeric(Main_Setting(7)) = 0 Then
				strTop = vbNullString
			Else
				strTop = "TOP " & CInt(Main_Setting(7))
			End If
			SQL = "SELECT " & strTop & " ChannelID,orders,ColorModes,FontModes,ChannelName,Caption,ChannelDir,StopChannel,IsHidden,BindDomain,DomainName,LinkTarget,ChannelType,ChannelUrl,IsHidden FROM [ECCMS_Channel] WHERE IsHidden = 0 Order By orders"
			Set Rs = Execute(SQL)
			If Rs.BOF And Rs.EOF Then
				strContent = ""
			Else
			i = 0
			TotalNumber = Rs.RecordCount
			Do While Not Rs.EOF
				i = i + 1
				If Rs("LinkTarget") <> 0 Then
					LinkTarget = " target=""_blank"""
				Else
					LinkTarget = ""
				End If
				HtmlContent = HtmlContent & Main_Setting(9)
				ChannelName = ReadFontMode(Rs("ChannelName"), Rs("ColorModes"), Rs("FontModes"))
				If Rs("ChannelType") < 2 Then
					ChannelUrl = InstallDir & Rs("ChannelDir")
				Else
					ChannelUrl = Rs("ChannelUrl")
				End If
				If Rs("StopChannel") <> 0 Then
					sCaption = "此频道暂时关闭,不能访问！"
				Else
					sCaption = Rs("Caption")
				End If
				strContent = "<a href=""" & ChannelUrl & """" & LinkTarget & " title=""" & sCaption & """ class=navmenu>" & ChannelName & "</a>"
				If i Mod CInt(Main_Setting(8)) = 0 Then strContent = strContent & "<br>"
				HtmlContent = Replace(HtmlContent, "{$ChannelMenu}", strContent)	
			Rs.MoveNext
			Loop
			End If
			Rs.Close: Set Rs = Nothing
			'Value = strContent
		End If
		'strContent = Value
		
		ChannelMenu = HtmlContent
	End Function
	'=============================================================
	'函数名：LoadSelectClass
	'作  用：载入缓存下拉分类列表
	'参  数：ChannelID   ----频道ID
	'返回值：下拉分类列表
	'=============================================================
	Public Function LoadSelectClass(ChannelID)
		Dim CacheSelClass, SQL, Rs1, i
		
		Name = "SelectClass" & ChannelID
		If ObjIsEmpty() Then
			SQL = "select ClassID,ClassName,depth,TurnLink,child,isallow from ECCMS_Classify where ChannelID = " & ChannelID & " order by rootid,orders"
			Set Rs1 = Execute(SQL)
			If Rs1.BOF And Rs1.EOF Then
				CacheSelClass = CacheSelClass & "<option>没有添加分类</option>"
			End If
			Do While Not Rs1.EOF
				If Rs1("TurnLink") <> 0 Then
					CacheSelClass = CacheSelClass & "<option value=""0"""
				Else
					If Rs1("depth") = 0 And Rs1("child") <> 0 Then
					
						'判断是否允许在该目录下添加内容
						if Rs1("isallow")=1 then
							CacheSelClass = CacheSelClass & "<option value=""" & Rs1("ClassID") & """"
						else
							CacheSelClass = CacheSelClass & "<option"

						end if
					Else
						'判断是否允许在该目录下添加内容
						if Rs1("isallow")=1 then
							CacheSelClass = CacheSelClass & "<option value=""" & Rs1("ClassID") & """"
						else
							CacheSelClass = CacheSelClass & "<option"

						end if

						'CacheSelClass = CacheSelClass & "<option value=""" & Rs1("ClassID") & """"
					End If
				End If
				
				
				CacheSelClass = CacheSelClass & " {ClassID=" & Rs1("ClassID") & "}>"
				If Rs1("depth") = 1 Then CacheSelClass = CacheSelClass & "　├ "
				If Rs1("depth") > 1 Then
					For i = 2 To Rs1("depth")
						CacheSelClass = CacheSelClass & "　"
					Next
					CacheSelClass = CacheSelClass & "　├ "
				End If
				CacheSelClass = CacheSelClass & Rs1("ClassName") & "</option>" & vbCrLf
				Rs1.MoveNext
			Loop
			Rs1.Close
			Set Rs1 = Nothing
			Value = CacheSelClass
		
			
		End If
		LoadSelectClass = Value

	End Function
	Public Function ClassJumpMenu(ChannelID)
		Dim CacheJumpMenu
		Dim Rs1
		Dim i
		Name = "ClassJumpMenu" & ChannelID
		If ObjIsEmpty() Then
			Set Rs1 = Execute("select ClassID,ChannelID,ClassName,depth,TurnLink,TurnLinkUrl from [ECCMS_Classify] where ChannelID = " & ChannelID & " order by rootid,orders")
			Do While Not Rs1.EOF
				If Rs1("TurnLink") <> 0 Then
					CacheJumpMenu = CacheJumpMenu & "<option value=""" & Rs1("TurnLinkUrl") & """ {ClassID=" & Rs1("classid") & "}"
				Else
					CacheJumpMenu = CacheJumpMenu & "<option value=""?ChannelID=" & Rs1("ChannelID") & "&sortid=" & Rs1("classid") & """ {ClassID=" & Rs1("classid") & "}"
				End If
				If Trim(Request("sortid")) <> "" Then
					If CLng(Request("sortid")) = Rs1("classid") Then CacheJumpMenu = CacheJumpMenu & " selected"
				End If
				CacheJumpMenu = CacheJumpMenu & ">"
				If Rs1("depth") = 1 Then CacheJumpMenu = CacheJumpMenu & "　├ "
				If Rs1("depth") > 1 Then
					For i = 2 To Rs1("depth")
						CacheJumpMenu = CacheJumpMenu & "　"
					Next
					CacheJumpMenu = CacheJumpMenu & "　├ "
				End If
				CacheJumpMenu = CacheJumpMenu & Rs1("ClassName") & "</option>" & vbCrLf
				Rs1.MoveNext
			Loop
			Rs1.Close
			Set Rs1 = Nothing
			Value = CacheJumpMenu
		End If
		ClassJumpMenu = Value
	End Function
	'================================================
	'函数名：GetRandomCode
	'作  用：系统分配随机代码
	'================================================
	Public Function GetRandomCode()
		Dim Ran, i, LengthNum
		
		LengthNum = 16
		GetRandomCode = ""
		For i = 1 To LengthNum
			Randomize
			Ran = CInt(Rnd * 2)
			Randomize
			If Ran = 0 Then
				Ran = CInt(Rnd * 25) + 97
				GetRandomCode = GetRandomCode & UCase(Chr(Ran))
			ElseIf Ran = 1 Then
				Ran = CInt(Rnd * 9)
				GetRandomCode = GetRandomCode & Ran
			ElseIf Ran = 2 Then
				Ran = CInt(Rnd * 25) + 97
				GetRandomCode = GetRandomCode & Chr(Ran)
			End If
		Next
	End Function
	'================================================
	' 函数名：CodeIsTrue
	' 作  用：检查验证码是否正确
	'================================================
	Public Function CodeIsTrue()
	    Dim CodeStr
	    CodeStr = Trim(Request("CodeStr"))
	    On Error Resume Next
	    If CStr(Session("GetCode")) = CStr(CodeStr) And CodeStr <> "" Then
		CodeIsTrue = True
		Session("GetCode") = Empty
	    Else
		CodeIsTrue = False
		Session("GetCode") = Empty
	    End If
	End Function
	Public Function CheckAdmin(ByVal Flag)
		Dim Rs, SQL
		Dim i, TempAdmin, AdminFlag, AdminGrade
		
		CheckAdmin = False
		On Error Resume Next
		SQL = "SELECT AdminGrade,Adminflag FROM ECCMS_Admin WHERE username='" & Replace(Session("AdminName"), "'", "''") & "' And password='" & Replace(Session("AdminPass"), "'", "''") & "' And isLock=0 And id=" & CLng(Session("AdminID"))
		Set Rs = Execute(SQL)
		If Rs.BOF And Rs.EOF Then
			CheckAdmin = False
			Set Rs = Nothing
			Exit Function
		Else
			AdminFlag = Rs("Adminflag")
			AdminGrade = Rs("AdminGrade")
		End If
		Rs.Close: Set Rs = Nothing
		If CInt(AdminGrade) = 999 Then
			CheckAdmin = True
			Exit Function
		Else
			If Trim(Flag) = "" Then Exit Function
			If AdminFlag = "" Then
				CheckAdmin = False
				Exit Function
			Else
				TempAdmin = Split(AdminFlag, ",")
				For i = 0 To UBound(TempAdmin)
					If Trim(LCase(TempAdmin(i))) = Trim(LCase(Flag)) Then
						CheckAdmin = True
						Exit For
					End If
				Next
			End If
		End If
	End Function
	'================================================
	'函数名：ReadAlpha
	'作  用：读取字符串的第一个字母
	'参  数：str   ----字符
	'返回值：返回第一个字母
	'================================================
	Public Function ReadAlpha(ByVal str)
		Dim strTemp
		If IsNull(str) Or Trim(str) = "" Then
			ReadAlpha = "A-9"
			Exit Function
		End If
		str = Trim(str)
		strTemp = 65536 + Asc(str)
		If (strTemp >= 45217 And strTemp <= 45252) Or (strTemp = 65601) Or (strTemp = 65633) Or (strTemp = 37083) Then
			ReadAlpha = "A-Z"
		ElseIf (strTemp >= 45253 And strTemp <= 45760) Or (strTemp = 65602) Or (strTemp = 65634) Or (strTemp = 39658) Then
			ReadAlpha = "B-Z"
		ElseIf (strTemp >= 45761 And strTemp <= 46317) Or (strTemp = 65603) Or (strTemp = 65635) Or (strTemp = 33405) Then
			ReadAlpha = "C-Z"
		ElseIf (strTemp >= 46318 And strTemp <= 46930) Or (strTemp >= 61884 And strTemp <= 61884) Or (strTemp = 65604) Or (strTemp >= 36820 And strTemp <= 38524) Or (strTemp = 65636) Then
			ReadAlpha = "D-Z"
		ElseIf (strTemp >= 46931 And strTemp <= 47009) Or (strTemp = 65605) Or (strTemp = 65637) Or (strTemp = 61513) Then
			ReadAlpha = "E-Z"
		ElseIf (strTemp >= 47010 And strTemp <= 47296) Or (strTemp = 65606) Or (strTemp = 65638) Or (strTemp = 61320) Or (strTemp = 63568) Or (strTemp = 36281) Then
			ReadAlpha = "F-Z"
		ElseIf (strTemp >= 47297 And strTemp <= 47613) Or (strTemp = 65607) Or (strTemp = 65639) Or (strTemp = 35949) Or (strTemp = 36089) Or (strTemp = 36694) Or (strTemp = 34808) Then
			ReadAlpha = "G-Z"
		ElseIf (strTemp >= 47614 And strTemp <= 48118) Or (strTemp >= 59112 And strTemp <= 59112) Or (strTemp = 65608) Or (strTemp = 65640) Then
			ReadAlpha = "H-Z"
		ElseIf (strTemp = 65641) Or (strTemp = 65609) Or (strTemp = 65641) Then
			ReadAlpha = "I-Z"
		ElseIf (strTemp >= 48119 And strTemp <= 49061 And strTemp <> 48739) Or (strTemp >= 62430 And strTemp <= 62430) Or (strTemp = 65610) Or (strTemp = 65642) Or (strTemp = 39048) Then
			ReadAlpha = "J-Z"
		ElseIf (strTemp >= 49062 And strTemp <= 49323) Or (strTemp = 65611) Or (strTemp = 65643) Then
			ReadAlpha = "K-Z"
		ElseIf (strTemp >= 49324 And strTemp <= 49895) Or (strTemp >= 58838 And strTemp <= 58838) Or (strTemp = 65612) Or (strTemp = 65644) Or (strTemp = 62418) Or (strTemp = 48739) Then
			ReadAlpha = "L-Z"
		ElseIf (strTemp >= 49896 And strTemp <= 50370) Or (strTemp = 65613) Or (strTemp = 65645) Then
			ReadAlpha = "M-Z"
		ElseIf (strTemp >= 50371 And strTemp <= 50613) Or (strTemp = 65614) Or (strTemp = 65646) Then
			ReadAlpha = "N-Z"
		ElseIf (strTemp >= 50614 And strTemp <= 50621) Or (strTemp = 65615) Or (strTemp = 65647) Then
			ReadAlpha = "O-Z"
		ElseIf (strTemp >= 50622 And strTemp <= 50905) Or (strTemp = 65616) Or (strTemp = 65648) Then
			ReadAlpha = "P-Z"
		ElseIf (strTemp >= 50906 And strTemp <= 51386) Or (strTemp >= 62659 And strTemp <= 63172) Or (strTemp = 65617) Or (strTemp = 65649) Then
			ReadAlpha = "Q-Z"
		ElseIf (strTemp >= 51387 And strTemp <= 51445) Or (strTemp = 65618) Or (strTemp = 65650) Then
			ReadAlpha = "R-Z"
		ElseIf (strTemp >= 51446 And strTemp <= 52217) Or (strTemp = 65619) Or (strTemp = 65651) Or (strTemp = 34009) Then
			ReadAlpha = "S-Z"
		ElseIf (strTemp >= 52218 And strTemp <= 52697) Or (strTemp = 65620) Or (strTemp = 65652) Then
			ReadAlpha = "T-Z"
		ElseIf (strTemp = 65621) Or (strTemp = 65653) Then
			ReadAlpha = "U-Z"
		ElseIf (strTemp = 65622) Or (strTemp = 65654) Then
			ReadAlpha = "V-Z"
		ElseIf (strTemp >= 52698 And strTemp <= 52979) Or (strTemp = 65623) Or (strTemp = 65655) Then
			ReadAlpha = "W-Z"
		ElseIf (strTemp >= 52980 And strTemp <= 53688) Or (strTemp = 65624) Or (strTemp = 65656) Then
			ReadAlpha = "X-Z"
		ElseIf (strTemp >= 53689 And strTemp <= 54480) Or (strTemp = 65625) Or (strTemp = 65657) Then
			ReadAlpha = "Y-Z"
		ElseIf (strTemp >= 54481 And strTemp <= 62383 And strTemp <> 59112 And strTemp <> 58838) Or (strTemp = 65626) Or (strTemp = 65658) Or (strTemp = 38395) Or (strTemp = 39783) Then
			ReadAlpha = "Z-Z"
		Else
			ReadAlpha = "A-9"
		End If
		If (strTemp >= 65633 And strTemp <= 65658) Or (strTemp >= 65601 And strTemp <= 65626) Then ReadAlpha = UCase(Left(str, 1))
		If (strTemp >= 65584 And strTemp <= 65593) Then ReadAlpha = "0-9"
	End Function
	'-- 修正文件路径
	Public Function CheckPath(ByVal sPath)
		sPath = Trim(sPath)
		If Right(sPath, 1) <> "\" And sPath <> "" Then
			sPath = sPath & "\"
		End If
		CheckPath = sPath
	End Function
	'-- 生成目录
	Public Function CreatPathEx(ByVal sPath)
		sPath = Replace(sPath, "/", "\")
		sPath = Replace(sPath, "\\", "\")
		On Error Resume Next
		
		Dim strHostPath,strPath
		Dim sPathItem,sTempPath
		Dim i,fso
		
		Set fso = Server.CreateObject(FSO_ScriptName)
		strHostPath = Server.MapPath("/")
		If InStr(sPath, ":") = 0 Then sPath = Server.MapPath(sPath)
		If fso.FolderExists(sPath) Or Len(sPath) < 3 Then
			CreationPath = True
			Exit Function
		End If
		
		strPath = Replace(sPath, strHostPath, vbNullString,1,-1,1)
		sPathItem = Split(strPath, "\")
		
		If InStr(LCase(sPath), LCase(strHostPath)) = 0 Then
			sTempPath = sPathItem(0)
		Else
			sTempPath = strHostPath
		End If
		
		For i = 1 To UBound(sPathItem)
			If sPathItem(i) <> "" Then
				sTempPath = sTempPath & "\" & sPathItem(i)
				If fso.FolderExists(sTempPath) = False Then
					fso.CreateFolder sTempPath
				End If
			End If
		Next
		Set fso = Nothing
		If Err.Number <> 0 Then Err.Clear
		CreatPathEx = True
	End Function
	'================================================
	'函数名：FilesDelete
	'作  用：FSO删除文件
	'参  数：filepath   ----文件路径
	'返回值：False  ----  True
	'================================================
	Public Function FileDelete(ByVal FilePath)
		On Error Resume Next
		FileDelete = False
		Dim fso
		Set fso = Server.CreateObject(FSO_ScriptName)
		If FilePath = "" Then Exit Function
		If InStr(FilePath, ":") = 0 Then FilePath = Server.MapPath(FilePath)
		If fso.FileExists(FilePath) Then
			fso.DeleteFile FilePath, True
			FileDelete = True
		End If
		Set fso = Nothing
		If Err.Number <> 0 Then Err.Clear
	End Function
	'================================================
	'函数名：FolderDelete
	'作  用：FSO删除目录
	'参  数：folderpath   ----目录路径
	'返回值：False  ----  True
	'================================================
	Public Function FolderDelete(ByVal FolderPath)
		FolderDelete = False
		On Error Resume Next
		Dim fso
		Set fso = Server.CreateObject(FSO_ScriptName)
		If FolderPath = "" Then Exit Function
		If InStr(FolderPath, ":") = 0 Then FolderPath = Server.MapPath(FolderPath)
		If fso.FolderExists(FolderPath) Then
			fso.DeleteFolder FolderPath, True
			FolderDelete = True
		End If
		Set fso = Nothing
		If Err.Number <> 0 Then Err.Clear
	End Function
	'================================================
	'函数名：CopyToFile
	'作  用：复制文件
	'参  数：SoureFile   ----原文件路径
	'        NewFile  ----目标文件路径
	'================================================
	Public Function CopyToFile(ByVal SoureFile, ByVal NewFile)
		On Error Resume Next
		If SoureFile = "" Then Exit Function
		If NewFile = "" Then Exit Function
		If InStr(SoureFile, ":") = 0 Then SoureFile = Server.MapPath(SoureFile)
		If InStr(NewFile, ":") = 0 Then NewFile = Server.MapPath(NewFile)
		Dim fso
		Set fso = Server.CreateObject(FSO_ScriptName)
		If fso.FileExists(SoureFile) Then
			fso.CopyFile SoureFile, NewFile
		End If
		Set fso = Nothing
		If Err.Number <> 0 Then Err.Clear
	End Function
	'================================================
	'函数名：CopyToFolder
	'作  用：复制文件夹
	'参  数：SoureFolder   ----原路径
	'        NewFolder  ----目标路径
	'================================================
	Public Function CopyToFolder(ByVal SoureFolder, ByVal NewFolder)
		On Error Resume Next
		If SoureFolder = "" Then Exit Function
		If NewFolder = "" Then Exit Function
		If InStr(SoureFolder, ":") = 0 Then SoureFolder = Server.MapPath(SoureFolder)
		If InStr(NewFolder, ":") = 0 Then NewFolder = Server.MapPath(NewFolder)
		Dim fso
		Set fso = Server.CreateObject(FSO_ScriptName)
		If fso.FolderExists(SoureFolder) Then
			fso.CopyFolder SoureFolder, NewFolder
		End If
		Set fso = Nothing
		If Err.Number <> 0 Then Err.Clear
	End Function
	'=============================================================
	'过程名：CreatedTextFile
	'作  用：创建文本文件
	'参  数：filename  ----文件名
	'        body  ----主要内容
	'=============================================================
	Public Function CreatedTextFile(ByVal FileName, ByVal body)
		On Error Resume Next
		If InStr(FileName, ":") = 0 Then FileName = Server.MapPath(FileName)
		Dim fso,f
		Set fso = Server.CreateObject(FSO_ScriptName)
		Set f = fso.CreateTextFile(FileName)
		f.WriteLine body
		f.Close
		Set f = Nothing
		Set fso = Nothing
		If Err.Number <> 0 Then Err.Clear
	End Function
	'================================================
	'函数名：Readfile
	'作  用：读取文件内容
	'参  数：fromPath   ----来源文件路径
	'================================================
	Public Function Readfile(ByVal fromPath)
		On Error Resume Next
		Dim strTemp,fso,f
		If InStr(fromPath, ":") = 0 Then fromPath = Server.MapPath(fromPath)
		Set fso = Server.CreateObject(FSO_ScriptName)
		If fso.FileExists(fromPath) Then
			Set f = fso.OpenTextFile(fromPath, 1, True)
			strTemp = f.ReadAll
			f.Close
			Set f = Nothing
		End If
		Set fso = Nothing
		Readfile = strTemp
		If Err.Number <> 0 Then Err.Clear
	End Function
	
	'================================================
	'函数名：CutMatchContent
	'作  用：截取相匹配的内容
	'参  数：Str   ----原字符串
	'        PatStr   ----符合条件字符
	'================================================
	Public Function CutMatchContent(ByVal str, ByVal start, ByVal last, ByVal Condition)
        
		Dim Match,s,re
		Dim FilterStr,MatchStr
		Dim strContent,ArrayFilter
		Dim i, n,bRepeat
		
		If Len(start) = 0 Or Len(last) = 0 Then Exit Function
		
		On Error Resume Next
		
		MatchStr = "(" & CorrectPattern(start) & ")(.+?)(" & CorrectPattern(last) & ")"

		Set re = New RegExp
		re.IgnoreCase = True
		re.Global = True
		re.Pattern = MatchStr
		Set s = re.Execute(str)
		n = 0
		For Each Match In s
			If n = 0 Then
				n = n + 1
				ReDim ArrayFilter(n)
				ArrayFilter(n) = Match
			Else
				bRepeat = False
				For i = 0 To UBound(ArrayFilter)
					If UCase(Match) = UCase(ArrayFilter(i)) Then
						bRepeat = True
						Exit For
					End If
				Next
				If bRepeat = False Then
					n = n + 1
					ReDim Preserve ArrayFilter(n)
					ArrayFilter(n) = Match
				End If
			End If
		Next
		
		Set s = Nothing
		Set re = Nothing
		
		If CBool(Condition) Then
			strContent = Join(ArrayFilter, "|||")
		Else
			strContent = Join(ArrayFilter, "|||")
			strContent = Replace(strContent, start, "")
			strContent = Replace(strContent, last, "")
		End If
		
		CutMatchContent = Replace(strContent, "|||", vbNullString, 1, 1)
	End Function
	
	Function CutFixContent(ByVal str, ByVal start, ByVal last, ByVal n)
		Dim strTemp
		On Error Resume Next
		If InStr(str, start) > 0 Then
			Select Case n
			Case 0  '左右都截取（都取前面）（去处关键字）
				strTemp = Right(str, Len(str) - InStr(str, start) - Len(start) + 1)
				strTemp = Left(strTemp, InStr(strTemp, last) - 1)
			Case Else  '左右都截取（都取前面）（保留关键字）
				strTemp = Right(str, Len(str) - InStr(str, start) + 1)
				strTemp = Left(strTemp, InStr(strTemp, last) + Len(last) - 1)
			End Select
		Else
			strTemp = ""
		End If
		CutFixContent = strTemp
	End Function
	Private Function CorrectPattern(ByVal str)
		str = Replace(str, "\", "\\")
		str = Replace(str, "~", "\~")
		str = Replace(str, "!", "\!")
		str = Replace(str, "@", "\@")
		str = Replace(str, "#", "\#")
		str = Replace(str, "%", "\%")
		str = Replace(str, "^", "\^")
		str = Replace(str, "&", "\&")
		str = Replace(str, "*", "\*")
		str = Replace(str, "(", "\(")
		str = Replace(str, ")", "\)")
		str = Replace(str, "-", "\-")
		str = Replace(str, "+", "\+")
		str = Replace(str, "[", "\[")
		str = Replace(str, "]", "\]")
		str = Replace(str, "<", "\<")
		str = Replace(str, ">", "\>")
		str = Replace(str, ".", "\.")
		str = Replace(str, "/", "\/")
		str = Replace(str, "?", "\?")
		str = Replace(str, "=", "\=")
		str = Replace(str, "|", "\|")
		str = Replace(str, "$", "\$")
		CorrectPattern = str
	End Function
	'=============================================================
	'函数名：UserGroupSetting
	'作  用：取用户级权限设置
	'参  数：gradeid   ----等级ID
	'=============================================================
	Public Function UserGroupSetting(ByVal gradeid)
		If Not IsNumeric(gradeid) Then
			gradeid = 0
		End If
		On Error Resume Next
		Dim Rs, SQL
		
		Name = "GroupSetting" & gradeid
		If ObjIsEmpty() Then
			SQL = "Select Groupname,GroupSet from [ECCMS_UserGroup] where Grades =" & gradeid
			Set Rs = Execute(SQL)
			If Rs.BOF And Rs.EOF Then
				UserGroupSetting = ""
				Set Rs = Nothing
				Exit Function
			End If
			Value = Rs("GroupSet") & Rs("Groupname")
			Set Rs = Nothing
		End If
		UserGroupSetting = Value
	End Function
	Private Sub LoadGroupSetting()
		Dim strGroupSetting
		Dim Rs, SQL
		Dim Grades
		Grades = CInt(membergrade)
		On Error Resume Next
		If Grades > 0 And memberid > 0 Then
			If binUserLong = False Then
				Set Rs = Execute("SELECT userid FROM [ECCMS_User] WHERE password='" & CheckRequest(memberpass, 45) & "' And UserGrade=" & Grades & " And UserLock=0 And  userid =" & memberid)
				If Rs.BOF And Rs.EOF Then
					Grades = 0
					Response.Cookies(Cookies_Name) = ""
					binUserLong = False
				Else
					binUserLong = True
				End If
				Set Rs = Nothing
			End If
		End If
		
		Name = "GroupSetting" & Grades
		If ObjIsEmpty() Then
			SQL = "Select Groupname,GroupSet from [ECCMS_UserGroup] where Grades =" & Grades
			Set Rs = Execute(SQL)
			If Rs.BOF And Rs.EOF Then
				Response.Cookies(Cookies_Name) = ""
				Set Rs = Nothing
				Exit Sub
			End If
			Value = Rs("GroupSet") & Rs("Groupname")
			Set Rs = Nothing
		End If
		blnGroupSetting = True
		strGroupSetting = Value
		arrGroupSetting = Split(strGroupSetting, "|||")
	End Sub
	Public Property Get GroupSetting(i)
		If Not blnGroupSetting Then LoadGroupSetting
		GroupSetting = arrGroupSetting(i)
	End Property
	Public Function ReadContent(ByVal strContent)
		On Error Resume Next
		Dim re, i
		Dim sContentKeyword, strKeyword
		
		Set re = New RegExp
		re.IgnoreCase = True
		re.Global = True
		'过滤危险脚本
		re.Pattern = "(<s+cript(.[^>]*)>)"
		strContent = re.Replace(strContent, "&lt;&#83cript$2&gt;")
		re.Pattern = "(<\/s+cript>)"
		strContent = re.Replace(strContent, "&lt;/&#83cript&gt;")
		re.Pattern = "(<body(.[^>]*)>)"
		strContent = re.Replace(strContent, "<body>")
		re.Pattern = "(<\!(.[^>]*)>)"
		strContent = re.Replace(strContent, "&lt;$2&gt;")
		re.Pattern = "(<\!)"
		strContent = re.Replace(strContent, "&lt;!")
		re.Pattern = "(-->)"
		strContent = re.Replace(strContent, "--&gt;")
		re.Pattern = "(javascript:)"
		strContent = re.Replace(strContent, "<i>javascript</i>:")
		
		If Trim(sContentKeyword) <> "" Then
			sContentKeyword = Split(ContentKeyword, "@@@")
			For i = 0 To UBound(sContentKeyword) - 1
				strKeyword = Split(sContentKeyword(i), "$$$")
				re.Pattern = "(" & strKeyword(0) & ")"
				strContent = re.Replace(strContent, "<a target=""_blank"" href=""" & strKeyword(1) & """ class=""wordstyle"">$1</a>")
			Next
		End If
		
		re.Pattern = "(\[i\])(.[^\[]*)(\[\/i\])"
		strContent = re.Replace(strContent, "<i>$2</i>")
		re.Pattern = "(\[u\])(.[^\[]*)(\[\/u\])"
		strContent = re.Replace(strContent, "<u>$2</u>")
		re.Pattern = "(\[b\])(.[^\[]*)(\[\/b\])"
		strContent = re.Replace(strContent, "<b>$2</b>")
		re.Pattern = "(\[fly\])(.*)(\[\/fly\])"
		strContent = re.Replace(strContent, "<marquee>$2</marquee>")

		re.Pattern = "\[size=([1-9])\](.[^\[]*)\[\/size\]"
		strContent = re.Replace(strContent, "<font size=$1>$2</font>")
		re.Pattern = "(\[center\])(.[^\[]*)(\[\/center\])"
		strContent = re.Replace(strContent, "<center>$2</center>")

		're.Pattern = "<IMG.[^>]*SRC(=| )(.[^>]*)>"
		'strContent = re.Replace(strContent, "<IMG SRC=$2 border=""0"">")
		're.Pattern = "<img(.[^>]*)>"

		'strContent = re.Replace(strContent, "<img$1 onload=""return imgzoom(this,550)"">")

		re.Pattern = "\[DIR=*([0-9]*),*([0-9]*)\](.[^\[]*)\[\/DIR]"
		strContent = re.Replace(strContent, "<embed src=$3 pluginspage=http://www.macromedia.com/shockwave/download/ width=$1 height=$2></embed>")
		re.Pattern = "\[QT=*([0-9]*),*([0-9]*)\](.[^\[]*)\[\/QT]"
		strContent = re.Replace(strContent, "<embed src=$3 width=$1 height=$2 autoplay=true loop=false controller=true playeveryframe=false cache=false scale=TOFIT bgcolor=#000000 kioskmode=false targetcache=false pluginspage=http://www.apple.com/quicktime/>")
		re.Pattern = "\[MP=*([0-9]*),*([0-9]*)\](.[^\[]*)\[\/MP]"
		strContent = re.Replace(strContent, "<embed type=application/x-oleobject codebase=http://activex.microsoft.com/activex/controls/mplayer/en/nsmp2inf.cab#Version=5,1,52,701 flename=mp src=$3  width=$1 height=$2></embed>")
		re.Pattern = "\[RM=*([0-9]*),*([0-9]*)\](.[^\[]*)\[\/RM]"
		strContent = re.Replace(strContent, "<OBJECT classid=clsid:CFCDAA03-8BE4-11cf-B84B-0020AFBBCCFA class=OBJECT id=RAOCX width=$1 height=$2><PARAM NAME=SRC VALUE=$3><PARAM NAME=CONSOLE VALUE=Clip1><PARAM NAME=CONTROLS VALUE=imagewindow><PARAM NAME=AUTOSTART VALUE=true></OBJECT><br><OBJECT classid=CLSID:CFCDAA03-8BE4-11CF-B84B-0020AFBBCCFA height=32 id=video2 width=$1><PARAM NAME=SRC VALUE=$3><PARAM NAME=AUTOSTART VALUE=-1><PARAM NAME=CONTROLS VALUE=controlpanel><PARAM NAME=CONSOLE VALUE=Clip1></OBJECT>")

		re.Pattern = "(\[FLASH\])(.[^\[]*)(\[\/FLASH\])"
		strContent = re.Replace(strContent, "<embed src=""$2"" quality=high pluginspage='http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash' type='application/x-shockwave-flash' width=500 height=400>$2</embed>")
		re.Pattern = "(\[FLASH=*([0-9]*),*([0-9]*)\])(.[^\[]*)(\[\/FLASH\])"
		strContent = re.Replace(strContent, "<embed src=""$4"" quality=high pluginspage='http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash' type='application/x-shockwave-flash' width=$2 height=$3>$4</embed>")
		re.Pattern = "\[UPLOAD=(gif|jpg|jpeg|bmp|png)\](.[^\[]*)(gif|jpg|jpeg|bmp|png)\[\/UPLOAD\]"
		strContent = re.Replace(strContent, "<br><A HREF=""$2$1"" TARGET=_blank><IMG SRC=""$2$1"" border=0 alt=按此在新窗口浏览图片 onload=""javascript:if(this.width>screen.width-333)this.width=screen.width-333""></A>")

		re.Pattern = "(\[UPLOAD=(.[^\[]*)\])(.[^\[]*)(\[\/UPLOAD\])"
		strContent = re.Replace(strContent, "<br><a href=""$3"">点击浏览该文件</a>")

		re.Pattern = "(\[URL\])(.[^\[]*)(\[\/URL\])"
		strContent = re.Replace(strContent, "<A HREF=""$2"" TARGET=_blank>$2</A>")
		re.Pattern = "(\[URL=(.[^\[]*)\])(.[^\[]*)(\[\/URL\])"
		strContent = re.Replace(strContent, "<A HREF=""$2"" TARGET=_blank>$3</A>")

		re.Pattern = "(\[EMAIL\])(.[^\[]*)(\[\/EMAIL\])"
		strContent = re.Replace(strContent, "<A HREF=""mailto:$2"">$2</A>")
		re.Pattern = "(\[EMAIL=(.[^\[]*)\])(.[^\[]*)(\[\/EMAIL\])"
		strContent = re.Replace(strContent, "<A HREF=""mailto:$2"" TARGET=_blank>$3</A>")

		re.Pattern = "(\[HTML\])(.[^\[]*)(\[\/HTML\])"
		strContent = re.Replace(strContent, "<table width='100%' border='0' cellspacing='0' cellpadding='6' bgcolor='#F6F6F6'><td><b>以下内容为程序代码:</b><br>$2</td></table>")
		re.Pattern = "(\[code\])(.[^\[]*)(\[\/code\])"
		strContent = re.Replace(strContent, "<table width='100%' border='0' cellspacing='0' cellpadding='6' bgcolor='#F6F6F6'><td><b>以下内容为程序代码:</b><br>$2</td></table>")

		re.Pattern = "(\[color=(.[^\[]*)\])(.[^\[]*)(\[\/color\])"
		strContent = re.Replace(strContent, "<font color=$2>$3</font>")
		re.Pattern = "(\[face=(.[^\[]*)\])(.[^\[]*)(\[\/face\])"
		strContent = re.Replace(strContent, "<font face=$2>$3</font>")
		re.Pattern = "\[align=(center|left|right)\](.*)\[\/align\]"
		strContent = re.Replace(strContent, "<div align=$1>$2</div>")

		re.Pattern = "(\[QUOTE\])(.*)(\[\/QUOTE\])"
		strContent = re.Replace(strContent, "<table cellpadding=0 cellspacing=0 border=1 WIDTH=94% bordercolor=#000000 bgcolor=#F2F8FF align=center  ><tr><td  ><table width=100% cellpadding=5 cellspacing=1 border=0><TR><TD BGCOLOR='#F6F6F6'>$2</table></table><br>")
		re.Pattern = "(\[move\])(.*)(\[\/move\])"
		strContent = re.Replace(strContent, "<MARQUEE scrollamount=3>$2</marquee>")
		re.Pattern = "\[GLOW=*([0-9]*),*(#*[a-z0-9]*),*([0-9]*)\](.[^\[]*)\[\/GLOW]"
		strContent = re.Replace(strContent, "<table width=$1 style=""filter:glow(color=$2, strength=$3)"">$4</table>")
		re.Pattern = "\[SHADOW=*([0-9]*),*(#*[a-z0-9]*),*([0-9]*)\](.[^\[]*)\[\/SHADOW]"
		strContent = re.Replace(strContent, "<table width=$1 style=""filter:shadow(color=$2, strength=$3)"">$4</table>")
		Set re = Nothing
		
		strContent = Replace(strContent, "[InstallDir_ChannelDir]", InstallDir & "/" & ChannelDir)
		strContent = Replace(strContent, "{", "&#123;")
		strContent = Replace(strContent, "}", "&#125;")
		strContent = Replace(strContent, "$", "&#36;")
		ReadContent = strContent
	End Function
	'================================================
	'函数名：CCh
	'        得到一位数字 N1 的汉字大写
	'        0 返回 ""
	'================================================
	Private Function CCh(N1)
		Select Case N1
			Case 0:CCh = "零"
			Case 1:CCh = "壹"
			Case 2:CCh = "贰"
			Case 3:CCh = "叁"
			Case 4:CCh = "肆"
			Case 5:CCh = "伍"
			Case 6:CCh = "陆"
			Case 7:CCh = "柒"
			Case 8:CCh = "捌"
			Case 9:CCh = "玖"
		End Select
	End Function
	'================================================
	'函数名：ChMoney
	'       得到数字 N1 的汉字大写
	'       最大为 千万位
	'================================================
	Public Function ChMoney(N1)
		Dim tMoney,lMoney,ST1,t1
		Dim tn,s1,s2,s3
		On Error Resume Next
		If N1 = 0 Then
			ChMoney = "零"
			Exit Function
		End If
		If N1 > 99999999 Then
			ChMoney = ""
			Exit Function
		End If
		If N1 < 0 Then
			ChMoney = "负" + ChMoney(Abs(N1))
			Exit Function
		End If
		tMoney = Trim(Cstr(N1))
		tn = InStr(tMoney, ".")  '小数位置
		s1 = ""

		If tn <> 0 Then
			ST1 = Right(tMoney, Len(tMoney) - tn)
			If ST1 <> "" Then
				t1 = Left(ST1, 1)
				ST1 = Right(ST1, Len(ST1) - 1)
				If t1 <> "0" Then
					s1 = s1 + CCh(eval(t1)) + "角"
				End If
				If ST1 <> "" Then
					t1 = Left(ST1, 1)
					s1 = s1 + CCh(eval(t1)) + "分"
				End If
			End If
			ST1 = Left(tMoney, tn - 1)
		Else
			ST1 = tMoney
		End If

		s2 = ""
		If ST1 <> "" Then
			t1 = Right(ST1, 1)
			ST1 = Left(ST1, Len(ST1) - 1)
			s2 = CCh(eval(t1)) + s2
		End If

		If ST1 <> "" Then
			t1 = Right(ST1, 1)
			ST1 = Left(ST1, Len(ST1) - 1)
			If t1 <> "0" Then
				s2 = CCh(eval(t1)) + "拾" + s2
			Else
				If Left(s2, 1) <> "零" Then s2 = "零" + s2
			End If
		End If

		If ST1 <> "" Then
			t1 = Right(ST1, 1)
			ST1 = Left(ST1, Len(ST1) - 1)
			If t1 <> "0" Then
				s2 = CCh(eval(t1)) + "佰" + s2
			Else
				If Left(s2, 1) <> "零" Then s2 = "零" + s2
			End If
		End If

		If ST1 <> "" Then
			t1 = Right(ST1, 1)
			ST1 = Left(ST1, Len(ST1) - 1)
			If t1 <> "0" Then
				s2 = CCh(eval(t1)) + "仟" + s2
			Else
				If Left(s2, 1) <> "零" Then s2 = "零" + s2
			End If
		End If

		s3 = ""
		If ST1 <> "" Then
			t1 = Right(ST1, 1)
			ST1 = Left(ST1, Len(ST1) - 1)
			s3 = CCh(eval(t1)) + s3
		End If

		If ST1 <> "" Then
			t1 = Right(ST1, 1)
			ST1 = Left(ST1, Len(ST1) - 1)
			If t1 <> "0" Then
				s3 = CCh(eval(t1)) + "拾" + s3
			Else
				If Left(s3, 1) <> "零" Then s3 = "零" + s3
			End If
		End If

		If ST1 <> "" Then
			t1 = Right(ST1, 1)
			ST1 = Left(ST1, Len(ST1) - 1)
			If t1 <> "0" Then
				s3 = CCh(eval(t1)) + "佰" + s3
			Else
				If Left(s3, 1) <> "零" Then s3 = "零" + s3
			End If
		End If

		If ST1 <> "" Then
			t1 = Right(ST1, 1)
			ST1 = Left(ST1, Len(ST1) - 1)
			If t1 <> "0" Then
				s3 = CCh(eval(t1)) + "仟" + s3
			End If
		End If
		
		If Right(s2, 1) = "零" Then s2 = Left(s2, Len(s2) - 1)
		If Len(s3) > 0 Then
		
			If Right(s3, 1) = "零" Then s3 = Left(s3, Len(s3) - 1)
			s3 = s3 & "万"
		End If

		ChMoney = IIf(s3 & s2 = "", s1, s3 & s2 & "元" & s1)

	End Function
	Function IIF(bTest, resultTRUE, resultFALSE)
		If bTest = True Then
			IIF = resultTRUE
		Else
			IIF = resultFALSE
		End If
	End Function
	
	Public Function CheckBadstr(str)
		If IsNull(str) Then
			CheckBadstr = vbNullString
			Exit Function
		End If
		str = Replace(str, Chr(0), vbNullString)
		str = Replace(str, Chr(34), vbNullString)
		str = Replace(str, "%", vbNullString)
		str = Replace(str, "@", vbNullString)
		str = Replace(str, "!", vbNullString)
		str = Replace(str, "^", vbNullString)
		str = Replace(str, "=", vbNullString)
		str = Replace(str, "--", vbNullString)
		str = Replace(str, "$", vbNullString)
		str = Replace(str, "'", vbNullString)
		str = Replace(str, ";", vbNullString)
		CheckBadstr = Trim(str)
	End Function

End Class
%>
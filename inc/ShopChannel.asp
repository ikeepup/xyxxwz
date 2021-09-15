<!--#include file="ubbcode.asp"-->
<%
Dim enchicms
Set enchicms = New ShopChannel_Cls

Class ShopChannel_Cls

	Private ChannelID, CreateHtml, IsShowFlush
	Private Rs,SQL,ChannelRootDir,HtmlContent,strIndexName,HtmlFilePath
	private shopid,classid,skinid,TradeExplain,TradeName,strInstallDir
	Private strFileDir, ParentID, strParent, strClassName, ChildStr, Child
	Private maxperpage, TotalNumber, TotalPageNum, CurrentPage, i,j
	private ForbidEssay,ListContent,HtmlTemplate,TempListContent
	Private FoundErr,PageType

	Public Property Let Channel(ChanID)
		ChannelID = ChanID
	End Property
	Public Property Let ShowFlush(para)
		IsShowFlush = para
	End Property

	Private Sub Class_Initialize()
		On Error Resume Next
		ChannelID = 3
		PageType = 0
		FoundErr = False
	End Sub

	Private Sub Class_Terminate()
		Set HTML = Nothing
	End Sub

	Public Sub MainChannel()
		enchiasp.ReadChannel(ChannelID)
		CreateHtml = CInt(enchiasp.IsCreateHtml)
		ChannelRootDir = enchiasp.InstallDir & enchiasp.ChannelDir
		strInstallDir = enchiasp.InstallDir
		strIndexName = "<a href=""" & ChannelRootDir & """>" & enchiasp.ChannelName & "</a>"
		
	End Sub
	'=================================================
	'��������BuildShopIndex
	'��  �ã���ʾ�̳���ҳ
	'=================================================
	Public Sub BuildShopIndex()
		On Error Resume Next
		LoadShopIndex
		If CreateHtml <> 0 Then
			'��ʾHTML
			Response.Write "<meta http-equiv=refresh content=0;url=index" & enchiasp.HtmlExtName & ">"
		Else
			Response.Write HtmlContent
		End If
	End Sub
	'=================================================
	'��������CreateShopIndex
	'��  �ã������̳���ҳ��HTML
	'=================================================
	Public Sub CreateShopIndex()
		On Error Resume Next
		LoadShopIndex
		Dim FilePath
		FilePath = enchiasp.InstallDir & enchiasp.ChannelDir & "index" & enchiasp.HtmlExtName
		enchiasp.CreatedTextFile FilePath, HtmlContent
		If IsShowFlush = 1 Then Response.Write "<li style=""font-size: 12px;"">����" & enchiasp.ModuleName & "��ҳHTML���... <a href=" & FilePath & " target=_blank>" & Server.MapPath(FilePath) & "</a></li>" & vbNewLine
		Response.Flush
	End Sub
	Private Sub LoadShopIndex()
		On Error Resume Next
		enchiasp.LoadTemplates ChannelID, 1, enchiasp.ChannelSkin
		HtmlContent = enchiasp.HtmlContent
		HtmlContent = Replace(HtmlContent,"{$ChannelRootDir}", ChannelRootDir)
		HtmlContent = Replace(HtmlContent,"{$InstallDir}", strInstallDir)
		HtmlContent = Replace(HtmlContent,"{$ChannelID}", ChannelID)
		HtmlContent = Replace(HtmlContent,"{$PageTitle}", enchiasp.ChannelName)
		HtmlContent = Replace(HtmlContent,"{$ModuleName}", enchiasp.ModuleName)
		HtmlContent = Replace(HtmlContent,"{$ShopIndex}", strIndexName)
		HtmlContent = Replace(HtmlContent, "{$IndexTitle}", strIndexName)
		HtmlContent = HTML.ReadAnnounceContent(HtmlContent,ChannelID)
		HtmlContent = ReadClassMenu(HtmlContent)
		HtmlContent = ReadClassMenubar(HtmlContent)
		HtmlContent = HTML.ReadArticlePic(HtmlContent)
		HtmlContent = HTML.ReadSoftPic(HtmlContent)
		HtmlContent = HTML.ReadSoftList(HtmlContent)
		HtmlContent = HTML.ReadArticleList(HtmlContent)
		HtmlContent = HTML.ReadShopList(HtmlContent)
		HtmlContent = HTML.ReadShopPic(HtmlContent)
		HtmlContent = HTML.ReadFlashList(HtmlContent)
		HtmlContent = HTML.ReadFlashPic(HtmlContent)
		HtmlContent = HTML.ReadFriendLink(HtmlContent)
		HtmlContent = HTML.ReadGuestList(HtmlContent)
		HtmlContent = HTML.ReadAnnounceList(HtmlContent)
		HtmlContent = HTML.ReadPopularArticle(HtmlContent)
		HtmlContent = HTML.ReadStatistic(HtmlContent)
		HtmlContent = HTML.ReadPopularFlash(HtmlContent)
		HtmlContent = HTML.ReadUserRank(HtmlContent)
		HtmlContent = Replace(HtmlContent, "{$SkinPath}", enchiasp.SkinPath)
		HtmlContent = Replace(HtmlContent,"{$InstallDir}", enchiasp.InstallDir)
		HtmlContent = HtmlContent
	End Sub

	'#############################\\ִ����Ʒ��Ϣ��ʼ//#############################
	'=================================================
	'��������BuildShopInfo
	'��  �ã���ʾ�̳���Ϣҳ��
	'=================================================
	Public Sub BuildShopInfo()
		If CreateHtml <> 0 Then
			Response.Redirect (ChannelRootDir & "index" & enchiasp.HtmlExtName)
			Exit Sub
		Else
			enchiasp.PreventInfuse
			shopid = enchiasp.ChkNumeric(Request("id"))
			Response.Write LoadShopInfo(shopid)
		End If
	End Sub

	Public Function LoadShopInfo(shopid)
		Dim PastPrice,NowPrice,strLinkSite
		Dim strProductImage,ProductImageUrl,arrImageSize
		
		On Error Resume Next
		
		SQL = "SELECT A.*,C.ClassName,C.ParentID,C.ParentStr,C.skinid,C.HtmlFileDir,C.ChildStr,C.UseHtml FROM [ECCMS_ShopList] A INNER JOIN [ECCMS_Classify] C On A.ClassID=C.ClassID WHERE A.ChannelID=" & ChannelID & " And A.isAccept > 0 And A.shopid=" & shopid
		Set Rs = enchiasp.Execute(SQL)
		If Rs.BOF And Rs.EOF Then
			LoadShopInfo = ""
			If CreateHtml = 0 Then
				Response.Write "<meta http-equiv=""refresh"" content=""2;url='/"">" & vbNewLine
				Response.Write "<p align=""center"" style=""font-size: 16px;color: red;"">�Բ��𣬸�ҳ�淢���˴����޷�����! ϵͳ������Զ�ת����վ��ҳ......</p>" & vbNewLine
			End If
			Set Rs = Nothing
			Exit Function
		End If

		If Rs("skinid") <> 0 Then
			skinid = Rs("skinid")
		Else
			skinid = enchiasp.ChannelSkin
		End If
		
		enchiasp.LoadTemplates ChannelID, 3, skinid
		TradeExplain = Rs("Explain")
		TradeExplain = UbbCode(TradeExplain)
		
		arrImageSize = Split(enchiasp.HtmlSetting(9), "|")
		If enchiasp.CheckNull(Rs("ProductImage")) Then
			ProductImageUrl = enchiasp.GetImageUrl(Rs("ProductImage"), enchiasp.ChannelDir)
			strProductImage = enchiasp.GetFlashAndPic(ProductImageUrl, CInt(arrImageSize(0)), CInt(arrImageSize(1)))
			strProductImage = "<a href='" & ChannelRootDir & "Previewimg.asp?shopid=" & shopid & "' title='" & Rs("TradeName") & "' target=_blank>" & strProductImage & "</a>"
		Else
			strProductImage = enchiasp.HtmlSetting(8)
		End If
		
		If enchiasp.CheckNull(Rs("LinkSite")) Then
			strLinkSite = Replace(enchiasp.HtmlSetting(11),"{$Linking}",Trim(Rs("LinkSite")))
		Else
			strLinkSite = Trim(enchiasp.HtmlSetting(10))
		End If
		
		HtmlContent = enchiasp.HtmlContent
		HtmlContent = Replace(HtmlContent, "{$InstallDir}", strInstallDir)
		HtmlContent = Replace(HtmlContent, "{$ChannelID}", ChannelID)
		HtmlContent = Replace(HtmlContent, "{$Marque}", enchiasp.ChkNull(Rs("Marque")))
		HtmlContent = Replace(HtmlContent, "{$Unit}", enchiasp.ChkNull(Rs("Unit")))
		HtmlContent = Replace(HtmlContent, "{$Supply}", enchiasp.ChkNull(Rs("Supply")))
		HtmlContent = Replace(HtmlContent, "{$Company}", enchiasp.ChkNull(Rs("Company")))
		HtmlContent = Replace(HtmlContent, "{$Best}", Rs("isBest"))
		HtmlContent = Replace(HtmlContent, "{$Star}", enchiasp.ChkNumeric(Rs("star")))
		HtmlContent = Replace(HtmlContent, "{$addTime}", Rs("addTime"))
		HtmlContent = Replace(HtmlContent, "{$Integral}", Rs("integral"))
		
		HtmlContent = Replace(HtmlContent, "{$LinkSite}", strLinkSite)
		HtmlContent = Replace(HtmlContent, "{$PastPrice}", FormatNumber(Rs("PastPrice"),2,-1))
		HtmlContent = Replace(HtmlContent, "{$NowPrice}", FormatNumber(Rs("NowPrice"),2,-1))
		HtmlContent = Replace(HtmlContent, "{$YinPrice}", FormatNumber(Rs("YinPrice"),2,-1))
		HtmlContent = Replace(HtmlContent, "{$OtherPrice}", FormatNumber(Rs("OtherPrice"),2,-1))
		HtmlContent = Replace(HtmlContent, "{$TradeExplain}", TradeExplain)
		HtmlContent = Replace(HtmlContent, "{$ProductImage}", strProductImage)
		
		If InStr(HtmlContent, "{$FrontProduct}") > 0 Then
			HtmlContent = Replace(HtmlContent, "{$FrontProduct}", FrontProduct(shopid))
		End If
		If InStr(HtmlContent, "{$NextProduct}") > 0 Then
			HtmlContent = Replace(HtmlContent, "{$NextProduct}", NextProduct(shopid))
		End If
		If InStr(HtmlContent, "{$ProductComment}") > 0 Then
			HtmlContent = Replace(HtmlContent, "{$ProductComment}", ProductComment(Rs("shopid")))
		End If
		
		HtmlContent = Replace(HtmlContent, "{$ModuleName}", enchiasp.ModuleName)
		HtmlContent = Replace(HtmlContent, "{$PageTitle}", Rs("TradeName"))
		HtmlContent = Replace(HtmlContent, "{$classid}", Rs("ClassID"))
		HtmlContent = Replace(HtmlContent, "{$TradeName}", Rs("TradeName"))
		HtmlContent = Replace(HtmlContent, "{$ShopID}", Rs("shopid"))
		HtmlContent = HTML.ReadCurrentStation(HtmlContent, ChannelID, Rs("ClassID"), Rs("ClassName"), Rs("ParentID"), Rs("ParentStr"), Rs("HtmlFileDir"))
		HtmlContent = HTML.ReadAnnounceContent(HtmlContent, ChannelID)
HtmlContent = HTML.ReadFriendLink(HtmlContent)
		HtmlContent = ReadClassMenubar(HtmlContent)
		HtmlContent = ReadClassMenu(HtmlContent)
		HtmlContent = HTML.ReadShopPic(HtmlContent)
		HtmlContent = HTML.ReadShopList(HtmlContent)
		HtmlContent = Replace(HtmlContent, "{$ChannelRootDir}", ChannelRootDir)
		HtmlContent = Replace(HtmlContent, "{$SkinPath}", enchiasp.SkinPath)
		HtmlContent = Replace(HtmlContent, "{$InstallDir}", strInstallDir)
		HtmlContent = Replace(HtmlContent, "{$ChannelID}", ChannelID)
		If CreateHtml <> 0 Then
			Call CreateShopInfo
		Else
			LoadShopInfo = HtmlContent
		End If
		Rs.Close: Set Rs = Nothing
	End Function

	'=================================================
	'��������CreateShopInfo
	'��  �ã������̳���ϢHTML
	'=================================================
	Private Sub CreateShopInfo()
		Dim HtmlFileName
		HtmlFilePath = enchiasp.InstallDir & enchiasp.ChannelDir & Rs("HtmlFileDir") & enchiasp.ShowDatePath(Rs("HtmlFileDate"), enchiasp.HtmlPath)
		enchiasp.CreatPathEx (HtmlFilePath)
		HtmlFileName = HtmlFilePath & enchiasp.ReadFileName(Rs("HtmlFileDate"), Rs("shopid"), enchiasp.HtmlExtName, enchiasp.HtmlPrefix, enchiasp.HtmlForm, CurrentPage)
		enchiasp.CreatedTextFile HtmlFileName, HtmlContent
		If IsShowFlush = 1 Then 
			Response.Write "<li style=""font-size: 12px;"">����" & enchiasp.ModuleName & "��ϢHTML���... <a href=" & HtmlFileName & " target=_blank>" & Server.MapPath(HtmlFileName) & "</a></li>" & vbNewLine
			Response.Flush
		End If
	End Sub
	'=================================================
	'��������FrontProduct
	'��  �ã���ʾ��һ��Ʒ
	'=================================================
	Private Function FrontProduct(shopid)
		Dim rsContext, SQL, HtmlFileUrl, HtmlFileName
		On Error Resume Next
		SQL = "select Top 1 A.shopid,A.ClassID,A.TradeName,A.HtmlFileDate,C.HtmlFileDir from [ECCMS_ShopList] A inner join [ECCMS_Classify] C On A.ClassID=C.ClassID where A.ChannelID=" & ChannelID & " And A.isAccept > 0 And A.shopid < " & shopid & " order by A.shopid desc"
		Set rsContext = enchiasp.Execute(SQL)
		If rsContext.EOF And rsContext.BOF Then
			FrontProduct = "�Ѿ�û����"
		Else
			If CreateHtml <> 0 Then
				HtmlFileUrl = ChannelRootDir & rsContext("HtmlFileDir") & enchiasp.ShowDatePath(rsContext("HtmlFileDate"), enchiasp.HtmlPath)
				HtmlFileName = enchiasp.ReadFileName(rsContext("HtmlFileDate"), rsContext("shopid"), enchiasp.HtmlExtName, enchiasp.HtmlPrefix, enchiasp.HtmlForm, "")
				FrontProduct = "<a href=" & HtmlFileUrl & HtmlFileName & ">" & rsContext("TradeName") & "</a>"
			Else
				FrontProduct = "<a href=?id=" & rsContext("shopid") & ">" & rsContext("TradeName") & "</a>"
			End If
		End If
		rsContext.Close
		Set rsContext = Nothing
	End Function
	'=================================================
	'��������NextProduct
	'��  �ã���ʾ��һ��Ʒ
	'=================================================
	Private Function NextProduct(shopid)
		Dim rsContext, SQL, HtmlFileUrl, HtmlFileName
		On Error Resume Next
		SQL = "select Top 1 A.shopid,A.ClassID,A.TradeName,A.HtmlFileDate,C.HtmlFileDir from [ECCMS_ShopList] A inner join [ECCMS_Classify] C On A.ClassID=C.ClassID where A.ChannelID=" & ChannelID & " And A.isAccept > 0 And A.shopid > " & shopid & " order by A.shopid asc"
		Set rsContext = enchiasp.Execute(SQL)
		If rsContext.EOF And rsContext.BOF Then
			NextProduct = "�Ѿ�û����"
		Else
			If CreateHtml <> 0 Then
				HtmlFileUrl = ChannelRootDir & rsContext("HtmlFileDir") & enchiasp.ShowDatePath(rsContext("HtmlFileDate"), enchiasp.HtmlPath)
				HtmlFileName = enchiasp.ReadFileName(rsContext("HtmlFileDate"), rsContext("shopid"), enchiasp.HtmlExtName, enchiasp.HtmlPrefix, enchiasp.HtmlForm, "")
				NextProduct = "<a href=" & HtmlFileUrl & HtmlFileName & ">" & rsContext("TradeName") & "</a>"
			Else
				NextProduct = "<a href=?id=" & rsContext("shopid") & ">" & rsContext("TradeName") & "</a>"
			End If
		End If
		rsContext.Close
		Set rsContext = Nothing
	End Function
	'#############################\\ִ����Ʒ�б�ʼ//#############################
	'=================================================
	'��������BuildShopList
	'��  �ã���ʾ�̳��б�ҳ��
	'=================================================
	Public Sub BuildShopList()
		If CreateHtml <> 0 Then
			Response.Redirect (ChannelRootDir & "index" & enchiasp.HtmlExtName)
			Exit Sub
		Else
			enchiasp.PreventInfuse
			If Not IsNumeric(Request("page")) And Len(Request("page")) <> 0 Then
				Response.Write ("�����ϵͳ����!����������")
				Response.End
			End If
			If Not IsEmpty(Request("page")) And Len(Request("page")) <> 0 Then
				CurrentPage = CLng(Request("page"))
			Else
				CurrentPage = 1
			End If
			classid = enchiasp.ChkNumeric(Request("classid"))
			Response.Write LoadShopList(ClassID, 1)
		End If
	End Sub
	'=================================================
	'��������LoadShopList
	'��  �ã������̳��б�
	'=================================================
	Public Function LoadShopList(clsid, n)
		On Error Resume Next
		Dim rsClass
		Dim HtmlFileName,maxparent,strMaxParent
		
		PageType = 1
		
		If Not IsNumeric(clsid) Then Exit Function
		Set rsClass = enchiasp.Execute("SELECT ClassID,ClassName,ChildStr,ParentID,ParentStr,Child,skinid,HtmlFileDir,UseHtml FROM [ECCMS_Classify] WHERE ChannelID = " & ChannelID & " And ClassID=" & clsid)
		If rsClass.BOF And rsClass.EOF Then
			If CreateHtml = 0 Then
				Response.Write "<meta http-equiv=""refresh"" content=""2;url='/"">" & vbNewLine
				Response.Write "<p align=""center"" style=""font-size: 12px;color: red;"">�Բ��𣬸�ҳ�淢���˴����޷�����! ϵͳ������Զ�ת����վ��ҳ......</p>" & vbNewLine
			End If
			Set rsClass = Nothing
			Exit Function
		Else
			strClassName = rsClass("ClassName")
			ClassID = rsClass("ClassID")
			ChildStr = rsClass("ChildStr")
			Child = rsClass("Child")
			strFileDir = rsClass("HtmlFileDir")
			ParentID = rsClass("ParentID")
			strParent = rsClass("ParentStr")
			If rsClass("skinid") <> 0 Then
				skinid = rsClass("skinid")
			Else
				skinid = CLng(enchiasp.ChannelSkin)
			End If
		End If
		rsClass.Close: Set rsClass = Nothing

		enchiasp.LoadTemplates ChannelID, 2, skinid
		HtmlFilePath = enchiasp.InstallDir & enchiasp.ChannelDir & strFileDir
		
		HtmlContent = Replace(enchiasp.HtmlContent, "|||@@@|||", "")
		HtmlContent = Replace(HtmlContent, "{$ChannelRootDir}", ChannelRootDir)
		HtmlContent = Replace(HtmlContent, "{$InstallDir}", strInstallDir)
		HtmlContent = Replace(HtmlContent, "{$ChannelID}", ChannelID)
		HtmlContent = Replace(HtmlContent, "{$ModuleName}", enchiasp.ModuleName)
		HtmlContent = Replace(HtmlContent, "{$ClassID}", ClassID)
		HtmlContent = Replace(HtmlContent, "{$PageTitle}", strClassName)
		HtmlContent = Replace(HtmlContent, "{$ShopIndex}", strIndexName)
		HtmlContent = Replace(HtmlContent, "{$IndexTitle}", strIndexName)
		HtmlContent = Replace(HtmlContent, "{$ClassName}", strClassName)

		ReplaceContent
		maxparent = enchiasp.ChkNumeric(enchiasp.HtmlSetting(5))
		maxperpage = CInt(enchiasp.HtmlSetting(1))
		
		If CLng(CurrentPage) = 0 Then CurrentPage = 1
		TotalNumber = enchiasp.Execute("SELECT COUNT(shopid) FROM ECCMS_ShopList WHERE ChannelID = " & ChannelID & " And isAccept > 0 And ClassID in (" & ChildStr & ")")(0)
		If maxparent > 0 And Child > 0 And TotalNumber > maxparent Then
			strMaxParent = " TOP " & maxparent
			TotalNumber = maxparent
		Else
			strMaxParent = ""
		End If
		TotalPageNum = CLng(TotalNumber / maxperpage)  '�õ���ҳ��
		If TotalPageNum < TotalNumber / maxperpage Then TotalPageNum = TotalPageNum + 1
		If CurrentPage < 1 Then CurrentPage = 1
		If CurrentPage > TotalPageNum Then CurrentPage = TotalPageNum
		
		Set Rs = CreateObject("ADODB.Recordset")
		SQL = "SELECT" & strMaxParent & " A.ShopID,A.ClassID,A.TradeName,A.Explain,A.PastPrice,A.NowPrice,A.star,A.ProductImage,A.addTime,A.AllHits,A.HtmlFileDate,A.isBest,C.ClassName,C.ParentID,C.ParentStr,C.skinid,C.HtmlFileDir,C.ChildStr,C.UseHtml FROM [ECCMS_ShopList] A INNER JOIN [ECCMS_Classify] C On A.ClassID=C.ClassID where A.ChannelID=" & ChannelID & " And A.isAccept > 0 And A.ClassID in (" & ChildStr & ") ORDER BY A.isTop DESC, A.addTime DESC ,A.shopid DESC"
		If isSqlDataBase = 1 Then
			Set Rs = enchiasp.Execute(SQL)
		Else
			Rs.Open SQL, Conn, 1, 1
		End If
		If Err.Number <> 0 Then Response.Write "SQL ��ѯ����"
		If Rs.BOF And Rs.EOF Then
			HtmlContent = Replace(HtmlContent, "{$ReadListPage}", "��û���ҵ��κ�" & enchiasp.ModuleName & "")
			HtmlContent = Replace(HtmlContent, enchiasp.CutFixContent(HtmlContent, "[ShowRepetend]", "[/ShowRepetend]", 1), "")
			If CreateHtml <> 0 Then
				enchiasp.CreatPathEx (HtmlFilePath)
				HtmlFileName = HtmlFilePath & ReadListPageName(ClassID, CurrentPage)
				enchiasp.CreatedTextFile HtmlFileName, HtmlContent
				If IsShowFlush = 1 Then 
					Response.Write "<li style=""font-size: 12px;"">����" & enchiasp.ModuleName & "�б�HTML���... <a href=" & HtmlFileName & " target=_blank>" & Server.MapPath(HtmlFileName) & "</a></li>" & vbNewLine
					Response.Flush
				End If
			End If
		Else
			TempListContent = enchiasp.CutFixContent(HtmlContent, "[ShowRepetend]", "[/ShowRepetend]", 1)
			If CreateHtml <> 0 Then
				Call LoadShopHtmlList(n)
			Else
				Call LoadShopAspList
			End If
		End If
		Rs.Close: Set Rs = Nothing
		
		LoadShopList = HtmlContent
	End Function
	'================================================
	'��������ReplaceContent
	'��  �ã��滻ģ������
	'================================================
	Private Sub ReplaceContent()
		HtmlContent = HTML.ReadCurrentStation(HtmlContent, ChannelID, ClassID, strClassName, ParentID, strParent, strFileDir)
		HtmlContent = HTML.ReadAnnounceContent(HtmlContent, ChannelID)
		HtmlContent = ReadClassMenubar(HtmlContent)
		HtmlContent = ReadClassMenu(HtmlContent)
		HtmlContent = HTML.ReadNewsPicAndText(HtmlContent)
		HtmlContent = HTML.ReadPopularArticle(HtmlContent)
		HtmlContent = HTML.ReadShopList(HtmlContent)
		HtmlContent = HTML.ReadShopPic(HtmlContent)
		HtmlContent = HTML.ReadStatistic(HtmlContent)
HtmlContent = HTML.ReadFriendLink(HtmlContent)
		HtmlContent = Replace(HtmlContent, "{$SkinPath}", enchiasp.SkinPath)
		HtmlContent = Replace(HtmlContent, "{$InstallDir}", strInstallDir)
		HtmlContent = Replace(HtmlContent, "{$ChannelID}", ChannelID)
	End Sub
	'================================================
	'��������LoadShopHtmlList
	'��  �ã�װ���̳��б�HTML
	'================================================
	Private Sub LoadShopHtmlList(n)
		Dim HtmlFileName
		Dim Perownum,ii,w
		
		Perownum = enchiasp.ChkNumeric(enchiasp.HtmlSetting(4))
		
		If IsNull(TempListContent) Then Exit Sub
		
		enchiasp.CreatPathEx (HtmlFilePath)
		For CurrentPage = n To TotalPageNum
			Rs.MoveFirst
			i = 0
			If TotalPageNum > 1 Then Rs.Move (CurrentPage - 1) * maxperpage
			ListContent = ""
			j = (CurrentPage - 1) * maxperpage + 1
			If Perownum > 1 Then 
				ListContent = enchiasp.HtmlSetting(6)
				w = FormatPercent(100 / Perownum / 100,0)
			End If
			
			Do While Not Rs.EOF And i < CInt(maxperpage)
				If Not Response.IsClientConnected Then Response.end
				
				If Perownum > 1 Then
					ListContent = ListContent & "<tr valign=""top"">" & vbCrLf
					For ii = 1 To Perownum
						ListContent = ListContent & "<td width=""" & w & """ class=""shoplistrow"">"
						If Not Rs.EOF Then
							Call LoadListDetail
							Rs.movenext
							i = i + 1
							j = j + 1
						End If
						ListContent = ListContent & "</td>" & vbCrLf
					Next
					ListContent = ListContent & "</tr>" & vbCrLf
				Else
					Call LoadListDetail
					Rs.MoveNext
					i = i + 1
					j = j + 1
				End If
				
				If i >= maxperpage Then Exit Do
			Loop
			
			Dim strHtmlFront, strHtmlPage
			
			strHtmlFront = enchiasp.HtmlPrefix & enchiasp.Supplemental(ClassID, 3) & "_"
			strHtmlPage = ShowHtmlPage(CurrentPage, TotalPageNum, TotalNumber, maxperpage, strHtmlFront, enchiasp.HtmlExtName, strClassName)
			HtmlTemplate = HtmlContent
			HtmlTemplate = Replace(HtmlTemplate, TempListContent, ListContent)
			HtmlTemplate = Replace(HtmlTemplate, "{$ReadListPage}", strHtmlPage)
			HtmlTemplate = Replace(HtmlTemplate, "[ShowRepetend]", "")
			HtmlTemplate = Replace(HtmlTemplate, "[/ShowRepetend]", "")
			'��ʼ�����ӷ����HTMLҳ
			HtmlFileName = HtmlFilePath & ReadListPageName(ClassID, CurrentPage)
			enchiasp.CreatedTextFile HtmlFileName, HtmlTemplate
			If IsShowFlush = 1 Then 
				Response.Write "<li style=""font-size: 12px;"">����" & enchiasp.ModuleName & "�б�HTML���... <a href=" & HtmlFileName & " target=_blank>" & Server.MapPath(HtmlFileName) & "</a></li>" & vbNewLine
				Response.Flush
			End If
		Next
		
	End Sub
	'================================================
	'��������LoadShopAspList
	'��  �ã�װ���̳��б�ASP
	'================================================
	Private Sub LoadShopAspList()
		Dim Perownum,ii,w
		
		If IsNull(TempListContent) Then Exit Sub
		
		Perownum = enchiasp.ChkNumeric(enchiasp.HtmlSetting(4))
		i = 0
		Rs.MoveFirst
		If TotalPageNum > 1 Then Rs.Move (CurrentPage - 1) * maxperpage
		ListContent = ""
		j = (CurrentPage - 1) * maxperpage + 1
		If Perownum > 1 Then 
			ListContent = enchiasp.HtmlSetting(6)
			w = FormatPercent(100 / Perownum / 100,0)
		End If
		
		Do While Not Rs.EOF And i < CInt(maxperpage)
			If Not Response.IsClientConnected Then Response.end
			
			If Perownum > 1 Then
				ListContent = ListContent & "<tr valign=""top"">" & vbCrLf
				For ii = 1 To Perownum
					ListContent = ListContent & "<td width=""" & w & """ class=""shoplistrow"">"
					If Not Rs.EOF Then
						Call LoadListDetail
						Rs.movenext
						i = i + 1
						j = j + 1
					End If
					ListContent = ListContent & "</td>" & vbCrLf
				Next
				ListContent = ListContent & "</tr>" & vbCrLf
			Else
				Call LoadListDetail
				Rs.MoveNext
				i = i + 1
				j = j + 1
			End If
			
			If i >= maxperpage Then Exit Do
		Loop
		If Perownum > 1 Then ListContent = ListContent & "</table>" & vbCrLf
		Dim strPagination
		strPagination = ShowListPage(CurrentPage, TotalPageNum, TotalNumber, maxperpage, ASPCurrentPage(PageType), strClassName)
		HtmlContent = Replace(HtmlContent, TempListContent, ListContent)
		HtmlContent = Replace(HtmlContent, "[ShowRepetend]", "")
		HtmlContent = Replace(HtmlContent, "[/ShowRepetend]", "")
		HtmlContent = Replace(HtmlContent, "{$ReadListPage}", strPagination)
	End Sub
	'================================================
	'��������LoadListDetail
	'��  �ã�װ���Ӽ�����б�ϸ��
	'================================================
	Private Sub LoadListDetail()
		Dim sTitle, sTopic, TradeName, ListStyle
		Dim ShopUrl, ShopTime, sClassName,strProductImage
		Dim ProductImageUrl, ProductImage,ProductIntro
		Dim strlen
		strlen = enchiasp.ChkNumeric(enchiasp.HtmlSetting(9))
		ListContent = ListContent & TempListContent
		If (i Mod 2) = 0 Then
			ListStyle = 1
		Else
			ListStyle = 2
		End If
		If strlen > 0 Then
			sTitle = enchiasp.GotTopic(Rs("TradeName"),strlen)
		Else
			sTitle = Rs("TradeName")
		End If
		On Error Resume Next
		If CInt(CreateHtml) <> 0 Then
			ShopUrl = ChannelRootDir & Rs("HtmlFileDir") & enchiasp.ShowDatePath(Rs("HtmlFileDate"), enchiasp.HtmlPath) & ReadPagination(0)
			sClassName = ChannelRootDir & Rs("HtmlFileDir")
		Else
			ShopUrl = ChannelRootDir & "show.asp?id=" & Rs("shopid")
			sClassName = ChannelRootDir & "list.asp?classid=" & Rs("ClassID")
		End If
		If Not IsNull(Rs("ProductImage")) Then
			strProductImage = Rs("ProductImage")
		End If
		ProductImageUrl = enchiasp.GetImageUrl(strProductImage, enchiasp.ChannelDir)
		ProductImage = enchiasp.GetFlashAndPic(ProductImageUrl, CInt(enchiasp.HtmlSetting(7)), CInt(enchiasp.HtmlSetting(8)))
		ProductImage = "<a href='" & ShopUrl & "' title='" & Rs("TradeName") & "'>" & ProductImage & "</a>"
		sClassName = "<a href='" & sClassName & "' title='" & Rs("ClassName") & "'>" & Rs("ClassName") & "</a>"
		TradeName = "<a href='" & ShopUrl & "' title='" & Rs("TradeName") & "' class=showtopic>" & sTitle & "</a>"

		ProductIntro = enchiasp.CutString(Rs("Explain"), CInt(enchiasp.HtmlSetting(3)))
		
		ShopTime = enchiasp.ShowDateTime(Rs("addTime"), CInt(enchiasp.HtmlSetting(2)))
		ListContent = Replace(ListContent, "{$ClassifyName}", sClassName)
		ListContent = Replace(ListContent, "{$TradeName}", TradeName)
		ListContent = Replace(ListContent, "{$ShopTopic}", sTitle)
		ListContent = Replace(ListContent, "{$ShopUrl}", ShopUrl)
		ListContent = Replace(ListContent, "{$ProductImage}", ProductImage)
		ListContent = Replace(ListContent, "{$ShopID}", Rs("shopid"))
		ListContent = Replace(ListContent, "{$ShopHits}", Rs("AllHits"))
		ListContent = Replace(ListContent, "{$Star}", Rs("star"))
		ListContent = Replace(ListContent, "{$ShopDateTime}", ShopTime)
		ListContent = Replace(ListContent, "{$ProductIntro}", ProductIntro)
		ListContent = Replace(ListContent, "{$ListStyle}", ListStyle)
		ListContent = Replace(ListContent, "{$PastPrice}", FormatNumber(Rs("PastPrice"),2,-1))
		ListContent = Replace(ListContent, "{$NowPrice}", FormatNumber(Rs("NowPrice"),2,-1))
		ListContent = Replace(ListContent, "{$IsTop}", Rs("IsTop"))
		ListContent = Replace(ListContent, "{$IsBest}", Rs("IsBest"))
		ListContent = Replace(ListContent, "{$Order}", j)
	End Sub
	'///---�̳��б����
	'///----------------------------------------------
	'///---���ﳵ���̿�ʼ
	'=================================================
	'��������BuildShopping
	'��  �ã����ﳵ
	'=================================================
	Public Sub BuildShopping()
		Dim strContent,Action
		Dim ProductIDList,ProductID,strProductID
		Dim strProductList,i,StyleList
		Dim Quantity,QuantityID,UnitPrice,TotalPrice
		Dim ShoppingHint,MaxProduct
		
		On Error Resume Next
		
		Action = LCase(enchiasp.CheckInfuse(Request("action"),8))
		skinid = CLng(enchiasp.ChannelSkin)
		enchiasp.LoadTemplates ChannelID, 6, skinid

		MaxProduct = enchiasp.ChkNumeric(enchiasp.HtmlSetting(1))
		If MaxProduct = 0 Then MaxProduct = 1
		'--����Ȩ������
		If CInt(enchiasp.GroupSetting(30)) = 0 Then
			Call OutAlertScript(enchiasp.CheckStr(enchiasp.HtmlSetting(8)))
			Exit Sub
		End If

		HtmlContent = enchiasp.HtmlContent
		HtmlContent = Replace(HtmlContent, "{$ChannelRootDir}", ChannelRootDir)
		HtmlContent = Replace(HtmlContent, "{$InstallDir}", strInstallDir)
		HtmlContent = Replace(HtmlContent, "{$ChannelID}", ChannelID)
		HtmlContent = Replace(HtmlContent, "{$ModuleName}", enchiasp.ModuleName)
		HtmlContent = Replace(HtmlContent, "{$ShopIndex}", strIndexName)
		HtmlContent = Replace(HtmlContent, "{$PageTitle}", "�ҵĹ��ﳵ")
		
		If enchiasp.CheckStr(action) = "ok" Then
			strProductList = enchiasp.CheckRequest(Request("ProductID"),0)
			Call ReformProduct(strProductList,MaxProduct)
		ElseIf enchiasp.CheckStr(action) = "del" Then
			Response.Cookies("ProductIDList") = ""
		ElseIf enchiasp.CheckStr(action) = "add" Then
		
			ProductID = enchiasp.ChkNumeric(Request("id"))
			If ProductID = 0 Then
				Call OutAlertScript(enchiasp.CheckStr(enchiasp.HtmlSetting(6)))
				Exit Sub
			End If
			
			ProductIDList = enchiasp.CheckRequest(Request.Cookies("ProductIDList")("ProductID"),0)
			Set Rs = CreateObject("ADODB.Recordset")
			SQL = "SELECT AllHits,DayHits,WeekHits,MonthHits,HitsTime fROM ECCMS_ShopList WHERE shopid = "& ProductID  
			Rs.Open SQL,Conn,1,3
			If Not (Rs.BOF And Rs.EOF) Then
				Rs("AllHits") = CCur(Rs("AllHits")) + 1
				If DateDiff("Ww", Rs("HitsTime"), Now) <= 0 Then
					Rs("WeekHits") = Rs("WeekHits") + 1
				Else
					Rs("WeekHits") = 1
				End If
				If DateDiff("M", Rs("HitsTime"), Now) <= 0 Then
					Rs("MonthHits") = Rs("MonthHits") + 1
				Else
					Rs("MonthHits") = 1
				End If
				If DateDiff("D", Rs("HitsTime"), Now) <= 0 Then
					Rs("DayHits") = Rs("DayHits") + 1
				Else
					Rs("DayHits") = 1
					Rs("HitsTime") = Now
				End If
				Rs.Update
			End If
			Rs.Close
			Set Rs = Nothing
			If Len(ProductIDList) = 0 Then
				ProductIDList = ProductID
			Else
			
				If CheckProductID(ProductID) Then
					ProductIDList = ProductID & "," & ProductIDList
				End If
				
			End If
			Call ReformProduct(ProductIDList,MaxProduct)
		End If
		
		strProductID = enchiasp.CheckRequest(Request.Cookies("ProductIDList")("ProductID"),0)
		If Len(strProductID) = 0 Then strProductID = 0

		strContent = enchiasp.HtmlSetting(2)
		If strProductID = "0" Then
			strContent = strContent & enchiasp.HtmlSetting(5)
			ShoppingHint = ""
		Else
			SQL = "SELECT TOP " & MaxProduct & " shopid,TradeName,NowPrice FROM [ECCMS_ShopList] WHERE ChannelID=" & ChannelID & " And isAccept > 0 And shopid in (" & strProductID & ")"
			Set Rs = enchiasp.Execute(SQL)
			If Rs.BOF And Rs.EOF Then
				strContent = strContent & enchiasp.HtmlSetting(5)
				ShoppingHint = ""
			Else
				ShoppingHint = enchiasp.HtmlSetting(7)
				i = 0
				Do While Not Rs.EOF
					If (i Mod 2) = 0 Then
						StyleList = 1
					Else
						StyleList = 2
					End If
					i = i + 1
					QuantityID = "Quantity_" & Rs("shopid")
					Quantity = enchiasp.ChkNumeric(Request(QuantityID))
					If Quantity = 0 Then Quantity = enchiasp.ChkNumeric(Request.Cookies("ProductIDList")(QuantityID))
					If Quantity = 0 Then Quantity = 1
					Response.Cookies("ProductIDList")(QuantityID) = Quantity
					UnitPrice = FormatNumber(Rs("NowPrice"), 2, -1)
					TotalPrice = FormatNumber(UnitPrice * Quantity, 2, -1)
					strContent = strContent & enchiasp.HtmlSetting(3)
					strContent = Replace(strContent, "{$ProductID}", Rs("shopid"))
					strContent = Replace(strContent, "{$TradeName}", Rs("TradeName"))
					strContent = Replace(strContent, "{$QuantityID}", QuantityID)
					strContent = Replace(strContent, "{$Quantity}", Quantity)
					strContent = Replace(strContent, "{$UnitPrice}", UnitPrice)
					strContent = Replace(strContent, "{$TotalPrice}", TotalPrice)
					strContent = Replace(strContent, "{$StyleList}", StyleList)
					strContent = Replace(strContent, "{$Ordered}", i)
					Rs.MoveNext
					
				Loop
			End If
			Set Rs = Nothing
		End If
		strContent = strContent & enchiasp.HtmlSetting(4)
		strContent = Replace(strContent, "{$ShoppingHint}", ShoppingHint)
		strContent = Replace(strContent, "{$MaxProduct}", MaxProduct)
HtmlContent = HTML.ReadFriendLink(HtmlContent)
		HtmlContent = Replace(HtmlContent, "{$PublicContent}", strContent)
		ReplaceString
		Response.Write HtmlContent
	End Sub
	'=================================================
	'��������CheckProductID
	'��  �ã��жϹ��ﳵ����Ʒ�Ƿ��ظ�
	'=================================================
	Private Function CheckProductID(ProductID)
		Dim strProductID,arrProductID,i
		On Error Resume Next
		
		ProductID = enchiasp.ChkNumeric(ProductID)

		If ProductID = 0 Then
			CheckProductID = False
			Exit Function
		End If

		strProductID = enchiasp.CheckRequest(Request.Cookies("ProductIDList")("ProductID"),0)
		
		If Len(strProductID) = 0 Then
			CheckProductID = True
			Exit Function
		End If
		arrProductID = Split(strProductID, ",")
		For i = 0 To UBound(arrProductID)
			If CLng(arrProductID(i)) = ProductID Then
				CheckProductID = False
				Exit Function
			End If
		Next
		CheckProductID = True
	End Function
	'=================================================
	'��������ReformProduct
	'��  �ã����鹺�ﳵ
	'=================================================
	Private Sub ReformProduct(strProductList,MaxProduct)
		Dim AllProductList
		Dim ArrayProduct(),arrProductList
		Dim i,n
		
		strProductList = Trim(strProductList)

		If Len(strProductList) = 0 Then
			Response.Cookies("ProductIDList") = ""
		Else
			arrProductList = Split(strProductList, ",")
			If UBound(arrProductList) > 0 Then
				n = 0
				For i = 0 To UBound(arrProductList)
					If i => MaxProduct Then Exit For
					If enchiasp.ChkNumeric(arrProductList(i)) > 0 Then
						ReDim Preserve ArrayProduct(n)
						ArrayProduct(n) = arrProductList(i)
						n = n + 1
					End If
				Next
				AllProductList = Join(ArrayProduct, ",")
			Else
				AllProductList = strProductList
			End If
			Response.Cookies("ProductIDList")("ProductID") = AllProductList
		End If
	End Sub
	'///---���ﳵ���̽���
	'///----------------------------------------------
	'///---�����ύ���̿�ʼ
	'=================================================
	'��������BuildPayment
	'��  �ã������ύ
	'=================================================
	Public Sub BuildPayment()
		Dim strContent,Action,OrderForm,ChineseMoney
		Dim ErrorMsg,Surcharge,Consignee,Company
		Dim Address,Phone,Postcode,Email,Oicq,Readme
		Dim ActualMoney,TotalMoney,Rebate,strRebate
		Dim strPayMode,PayMode,curdate,sRnd,userid
		Dim PayDone,UserName,UserGrade,strProductID
		Dim BuyCode
		
		On Error Resume Next

		strProductID = enchiasp.CheckRequest(Request.Cookies("ProductIDList")("ProductID"),0)
		userid = Clng(enchiasp.memberid)
		UserName = enchiasp.CheckRequest(enchiasp.membername,45)
		UserGrade = CInt(enchiasp.membergrade)
		Action = LCase(enchiasp.CheckInfuse(Request("action"),8))
		skinid = CLng(enchiasp.ChannelSkin)
		enchiasp.LoadTemplates ChannelID, 6, skinid

		'--����Ȩ������
		If userid = 0 Or UserName = "" Then
			Call OutAlertScript("�Բ���ֻ��ע���Ա����ʹ�ù��ﳵ����,����ע����½��")
			Exit Sub
		End If

		If CInt(enchiasp.GroupSetting(30)) = 0 Then
			Call OutAlertScript(enchiasp.CheckStr(enchiasp.HtmlSetting(8)))
			Exit Sub
		End If
		Rebate = CCur(enchiasp.GroupSetting(28))
		If Rebate > 0 And Rebate < 10 Then
			strRebate = Rebate & " ��"
		Else
			strRebate = "��"
			Rebate = 0
		End If
		If Len(strProductID) = 0 Then
			FoundErr = True
			ErrorMsg = "��Ĺ��ﳵΪ�գ���ѡ����Ʒ�����ύ��"
			Response.Cookies("ProductIDList") = ""
		End If
		If Trim(Action) = "ok" Or Trim(Action) = "pay" Then
			If Trim(Request.Form("surcharge")) = "" Then
				FoundErr = True
				ErrorMsg = "��ѡ�����ͷ�ʽ��"
			Else
				Surcharge = enchiasp.CheckNumeric(Request.Form("surcharge"))
			End if
			If Trim(Request.Form("consignee")) = "" Then
				FoundErr = True
				ErrorMsg = "�ջ������Ʋ���Ϊ�գ�"
			Else
				Consignee = enchiasp.CheckInfuse(Request.Form("consignee"),45)
			End if
			If Trim(Request.Form("company")) = "" Then
				Company = Trim(Request.Form("company"))
			Else
				Company = enchiasp.CheckInfuse(Request.Form("company"),180)
			End if
			If Trim(Request.Form("address")) = "" Then
				FoundErr = True
				ErrorMsg = "�ջ��˵�ַ����Ϊ�գ�"
			Else
				Address = enchiasp.CheckInfuse(Request.Form("address"),180)
			End if
			If Trim(Request.Form("phone")) = "" Then
				FoundErr = True
				ErrorMsg = "�ջ��˵绰����Ϊ�գ�"
			Else
				Phone = enchiasp.CheckInfuse(Request.Form("phone"),35)
			End if
			If Trim(Request.Form("postcode")) = "" Then
				FoundErr = True
				ErrorMsg = "�ջ����������벻��Ϊ�գ�"
			Else
				Postcode = enchiasp.CheckInfuse(Request.Form("postcode"),35)
			End if
			If Not CheckEmail(Request.Form("email")) Then
				FoundErr = True
				ErrorMsg = "�ջ���Email�������"
			Else
				Email = enchiasp.CheckInfuse(Request.Form("email"),45)
			End if
			If Trim(Request.Form("oicq")) = "" Then
				Oicq = Trim(Request.Form("oicq"))
			Else
				Oicq = enchiasp.CheckInfuse(Request.Form("oicq"),30)
			End if
			If Trim(Request.Form("Readme")) = "" Then
				Readme = Trim(Request.Form("Readme"))
			Else
				Readme = enchiasp.CheckRequest(Request.Form("Readme"),220)
			End if
			If Trim(Request.Form("OrderID")) = "" Then
				FoundErr = True
				ErrorMsg = "���׶����Ų���Ϊ�գ�"
			Else
				OrderForm = enchiasp.CheckInfuse(Request.Form("OrderID"),45)
			End if
		End If
		
		HtmlContent = enchiasp.HtmlContent
		HtmlContent = Replace(HtmlContent, "{$ChannelRootDir}", ChannelRootDir)
		HtmlContent = Replace(HtmlContent, "{$InstallDir}", strInstallDir)
		HtmlContent = Replace(HtmlContent, "{$ChannelID}", ChannelID)
		HtmlContent = Replace(HtmlContent, "{$ModuleName}", enchiasp.ModuleName)
		HtmlContent = Replace(HtmlContent, "{$ShopIndex}", strIndexName)
		
		Select Case Trim(Action)
		Case "ok" '--����ȷ��
			HtmlContent = Replace(HtmlContent, "{$PageTitle}", "����ȷ��")
			TotalMoney = CountTotalMoney

			If Rebate = 0 Then
				ActualMoney = TotalMoney + Surcharge
			Else
				ActualMoney = enchiasp.RebateMoney(TotalMoney,Rebate) + Surcharge
			End If
			If TotalMoney = 0 Then
				ErrorMsg ="�Ƿ���������ȡ���׶���󣡣���"
				Founderr = True
			End If
			If ActualMoney = 0 Then
				ErrorMsg ="�Ƿ���������ȡ���׶���󣡣���"
				Founderr = True
			End If
			
			Surcharge = FormatNumber(Surcharge,2,-1)
			TotalMoney = FormatNumber(TotalMoney,2,-1)
			ActualMoney = FormatNumber(ActualMoney,2,-1)
			ChineseMoney = enchiasp.ChMoney(CCur(ActualMoney))
		
			If FoundErr = False Then
				strContent = enchiasp.HtmlSetting(11)
				strContent = Replace(strContent, "{$Surcharge}", Surcharge)
				strContent = Replace(strContent, "{$Consignee}", Consignee)
				strContent = Replace(strContent, "{$Company}", Company)
				strContent = Replace(strContent, "{$Address}", Address)
				strContent = Replace(strContent, "{$Phone}", Phone)
				strContent = Replace(strContent, "{$Postcode}", Postcode)
				strContent = Replace(strContent, "{$Email}", Email)
				strContent = Replace(strContent, "{$Oicq}", Oicq)
				strContent = Replace(strContent, "{$Readme}", Readme)
				strContent = Replace(strContent, "{$ActualMoney}", ActualMoney)
				strContent = Replace(strContent, "{$TotalMoney}", TotalMoney)
				strContent = Replace(strContent, "{$Rebate}", strRebate)
				strContent = Replace(strContent, "{$Discount}", Rebate)
				strContent = Replace(strContent, "{$OrderID}", OrderForm)
				strContent = Replace(strContent, "{$ChineseMoney}", ChineseMoney)
				If CInt(enchiasp.GroupSetting(37)) = 0 Then
					strContent = Replace(strContent, "{$CodeStr}", 9999)
					strContent = Replace(strContent, "{$CodeSetting}", " style=""display:none""")
				Else
					strContent = Replace(strContent, "{$CodeStr}", "")
					strContent = Replace(strContent, "{$CodeSetting}", "")
				End If
			Else
				strContent = enchiasp.HtmlSetting(14)
				strContent = Replace(strContent, "{$ErrorMsg}", ErrorMsg)
			End If
		Case "pay"  '--����ȷ�ϳɹ�
			HtmlContent = Replace(HtmlContent, "{$PageTitle}", "�����ύ�ɹ�")
			PayMode = enchiasp.ChkNumeric(Request("PayMode"))
			TotalMoney = CCur(CountTotalMoney)
			
			If TotalMoney = 0 Then
				ErrorMsg ="�Ƿ���������ȡ���׶���󣡣���"
				Founderr = True
			End If
			
			If Rebate = 0 Then
				ActualMoney = CCur(TotalMoney + Surcharge)
			Else
				ActualMoney = enchiasp.RebateMoney(TotalMoney,Rebate) + Surcharge
			End If
			
			ActualMoney = CCur(FormatNumber(ActualMoney))
			ChineseMoney = enchiasp.ChMoney(ActualMoney)
			PayDone = 0

			If ActualMoney = 0 Then
				ErrorMsg ="�Ƿ���������ȡ���׶���󣡣���"
				Founderr = True
			End If
			
			If CInt(enchiasp.GroupSetting(37)) = 1 Then
				If Not enchiasp.CodeIsTrue() Then
					ErrorMsg ="��֤��У��ʧ�ܣ��뷵��ˢ��ҳ�����ԡ�"
					Founderr = True
				End If
			End If
			Session("GetCode") = ""
			
			Select Case PayMode
			Case 0
				strPayMode = "���л��"
			Case 1
				strPayMode = "����֧��"
			Case 2
				strPayMode = "վ��֧��"
				If userid > 0 Then
					Set Rs = enchiasp.Execute("SELECT userid,BuyCode,usermoney FROM ECCMS_User WHERE UserName='"& UserName &"' And UserGrade="& UserGrade &" And userid=" & userid)
					If Rs.BOF And Rs.EOF Then
						ErrorMsg ="�Ƿ�����������"
						Founderr = True
					Else
						
						BuyCode = md5(Trim(Request.Form("BuyCode")), False)
						'--�����û����
						If Rs("usermoney") < ActualMoney Then
							ErrorMsg ="�Բ�������ʻ����㣬��ʹ��������ʽ֧����"
							Founderr = True
						Else	'--�����û�վ��֧������
							If Trim(Rs("BuyCode")) <> BuyCode And Trim(Rs("BuyCode")) <> "" Then
								ErrorMsg ="�Բ���վ��֧����������뷵������ˢ��ҳ�����ԡ�"
								Founderr = True
							Else
								PayDone = 1
								enchiasp.Execute ("UPDATE ECCMS_User SET usermoney=usermoney-" & ActualMoney & ",prepaid=prepaid+" & ActualMoney & " WHERE userid=" & Rs("userid"))
							End If
						End If
					End If
					Set Rs = Nothing
				Else
					ErrorMsg ="�㲻�ǻ�Ա������ʹ��վ��֧��������"
					Founderr = True
				End If
			Case 3
				strPayMode = "�ʾֻ��"
			Case Else
				strPayMode = "�������"
			End Select
			
			strProductID = enchiasp.CheckRequest(Request.Cookies("ProductIDList")("ProductID"),0)
			If Len(strProductID) = 0 Then
				ErrorMsg ="�����������Ҳ�����ض�����Ϣ������"
				Founderr = True
			End If
			Set Rs = enchiasp.Execute("SELECT id FROM ECCMS_OrderForm WHERE OrderID='"& OrderForm &"'")
			If Not (Rs.BOF And Rs.EOF) Then
				ErrorMsg ="���Ѿ��ύ�˱����벻Ҫ�ظ��ύ������"
				Founderr = True
			End If
			Set Rs = Nothing
			
			If FoundErr = False Then
				Set Rs = CreateObject("ADODB.Recordset")
				SQL = "SELECT * FROM ECCMS_OrderForm WHERE (id is null)"
				Rs.Open SQL,Conn,1,3
				Rs.AddNew
					If userid > 0 Then
						Rs("userid") = userid
						Rs("username") = username
					Else
						Rs("userid") = 0
						Rs("username") = "�����û�"
					End If
					Rs("ProductID") = enchiasp.CheckStr(strProductID)
					Rs("OrderID") = enchiasp.CheckStr(OrderForm)
					Rs("Surcharge") = Surcharge
					Rs("totalmoney") = ActualMoney
					Rs("Consignee") = Consignee
					Rs("Company") = Company
					Rs("Address") = Address
					Rs("postcode") = postcode
					Rs("phone") = phone
					Rs("Email") = Email
					Rs("oicq") = oicq
					Rs("Readme") = Readme
					Rs("Paymode") = strPayMode
					Rs("addTime") = Now()
					Rs("invoice") = enchiasp.ChkNumeric(Request.Form("invoice"))
					Rs("finish") = 0
					Rs("Cancel") = 0
					Rs("PayDone") = PayDone
				Rs.Update
				Rs.Close:Set Rs = Nothing
				
				Set Rs = CreateObject("ADODB.Recordset")
				Rs.Open "SELECT id FROM ECCMS_OrderForm WHERE OrderID='"& OrderForm &"' ORDER BY id DESC", Conn, 1, 1
				Call AddBuyProduct(Rs("id"))
				Rs.Close:Set Rs = Nothing
				
				Dim wp,arrChinaeBank
				Dim strPlatform,SubmitCode
				strPlatform = ""
				'--�Ƿ������֧��
				If CInt(enchiasp.StopBankPay) > 0 And PayMode <> 2 Then
					arrChinaeBank = Split(enchiasp.ChinaeBank, "|||")
					SubmitCode = enchiasp.HtmlSetting(15)
					
					Set wp = New WebPayment_Cls
					wp.PayPlatform = CInt(enchiasp.StopBankPay)
					wp.submitvalue = SubmitCode
					wp.Paymentid = Trim(arrChinaeBank(0))
					wp.Paymentkey = Trim(arrChinaeBank(1))
					wp.Percent = enchiasp.CheckNumeric(arrChinaeBank(2))
					If LCase(Left(ChannelRootDir,7)) = "http://" Then
						wp.Returnurl = ChannelRootDir & "receive.asp"
					Else
						wp.Returnurl = enchiasp.GetSiteUrl & ChannelRootDir &"receive.asp"
					End If
					wp.Orderid = OrderForm
					wp.Paymoney = ActualMoney
					If Trim(Readme) = "" Then
						wp.Comment = "���Ϲ���"
					Else
						wp.Comment = Readme
					End If
					wp.Consignee = Consignee
					wp.Consigner = Consignee
					wp.Address = Address
					wp.Postcode = Postcode
					wp.Email = Email
					wp.Telephone = Phone
					strPlatform = wp.ShowPayment
					Set wp = Nothing
				End If
				strContent = enchiasp.HtmlSetting(12)
				strContent = Replace(strContent, "{$Surcharge}", FormatNumber(Surcharge,2,-1))
				strContent = Replace(strContent, "{$ActualMoney}", FormatNumber(ActualMoney,2,-1))
				strContent = Replace(strContent, "{$TotalMoney}", FormatNumber(TotalMoney,2,-1))
				strContent = Replace(strContent, "{$Rebate}", strRebate)
				strContent = Replace(strContent, "{$OrderID}", OrderForm)
				strContent = Replace(strContent, "{$ChineseMoney}", ChineseMoney)
				strContent = Replace(strContent, "{$WebPlatform}", strPlatform)
				If PayMode = 2 Then
					strContent = Replace(strContent, "{$SitePayInfo}", "��ϲ����վ��֧���ɹ������ν�������ɡ�")
				Else
					strContent = Replace(strContent, "{$SitePayInfo}", "�����ύ��ɣ�ֻ�и���ɹ��󣬱��ν��ײ�����ɡ�����ѡ������ʵʱ֧�������������ʽ��")
				End If
				Response.Cookies("ProductIDList") = ""
			Else
				strContent = enchiasp.HtmlSetting(14)
				strContent = Replace(strContent, "{$ErrorMsg}", ErrorMsg)
			End If
		Case Else
			'--�ύ����
			HtmlContent = Replace(HtmlContent, "{$PageTitle}", "�����ύ")
			If enchiasp.memberid > 0 Then
				Set Rs = enchiasp.Execute("SELECT userid,UserName,TrueName,usermail,phone,oicq,postcode,address FROM ECCMS_User WHERE UserName='"& UserName &"' And userid=" & userid)
				If Not (Rs.BOF And Rs.EOF) Then
					Consignee = Rs("TrueName")
					Address = Rs("address")
					Phone = Rs("phone")
					Postcode = Rs("postcode")
					Email = Rs("usermail")
					Oicq = Rs("oicq")
				End If
				Set Rs = Nothing
			End If
			If FoundErr = False Then
				Randomize
				sRnd = Int(9000 * Rnd) + 1000
				curdate=now()                                               
				OrderForm = Year(curdate) & Month(curdate) & Day(curdate) &"-"& sRnd &"-"& Hour(curdate) & Minute(curdate) & Second(curdate)

				strContent = enchiasp.HtmlSetting(9)
				strContent = strContent & enchiasp.HtmlSetting(10)
				strContent = Replace(strContent, "{$Consignee}", Consignee)
				strContent = Replace(strContent, "{$Address}", Address)
				strContent = Replace(strContent, "{$Phone}", Phone)
				strContent = Replace(strContent, "{$Postcode}", Postcode)
				strContent = Replace(strContent, "{$Email}", Email)
				strContent = Replace(strContent, "{$Oicq}", Oicq)
				strContent = Replace(strContent, "{$OrderID}", OrderForm)
			Else
				strContent = enchiasp.HtmlSetting(14)
				strContent = Replace(strContent, "{$ErrorMsg}", ErrorMsg)
				strContent = Replace(strContent, "{$DateTime}", Now())
			End If
		End Select
		
		HtmlContent = Replace(HtmlContent, "{$PublicContent}", strContent)
		ReplaceString
		Response.Write HtmlContent
	End Sub
	'=================================================
	'��������AddBuyProduct
	'��  �ã���ӹ�����Ʒ
	'=================================================
	Private Sub AddBuyProduct(sid)
		Dim strProductID,QuantityID
		Dim Quantity,UnitPrice,TotalPrice
		
		On Error Resume Next
		sid = CLng(sid)
		strProductID = enchiasp.CheckRequest(Request.Cookies("ProductIDList")("ProductID"),0)
		If strProductID = "" Then Exit Sub
		If Founderr = True Then Exit Sub
		SQL = "SELECT shopid,TradeName,NowPrice FROM [ECCMS_ShopList] WHERE ChannelID=" & ChannelID & " And isAccept > 0 And shopid in (" & strProductID & ")"
		Set Rs = enchiasp.Execute(SQL)
		If Not (Rs.BOF And Rs.EOF) Then
			Do While Not Rs.EOF
				QuantityID = "Quantity_" & Rs("shopid")
				Quantity = enchiasp.ChkNumeric(Request.Cookies("ProductIDList")(QuantityID))

				If Quantity = 0 Then Quantity = 1
				
				UnitPrice = Rs("NowPrice")
				TotalPrice = UnitPrice * Quantity
				SQL = "INSERT INTO ECCMS_buy (orderid,userid,shopid,TradeName,Amount,Price,totalmoney) VALUES ("& sid &","& enchiasp.memberid &","& Rs("shopid") &",'"& enchiasp.CheckStr(Rs("TradeName")) &"',"& Quantity &","& UnitPrice &","& TotalPrice &")"
				enchiasp.Execute(SQL)
				Rs.MoveNext
			Loop
		End If
		Set Rs = Nothing
	End Sub
	'=================================================
	'��������CountTotalMoney
	'��  �ã�ͳ���ܽ��
	'=================================================
	Public Function CountTotalMoney()
		Dim strProductID,QuantityID
		Dim Quantity,UnitPrice,TotalPrice
		CountTotalMoney = 0
		On Error Resume Next
		
		strProductID = enchiasp.CheckRequest(Request.Cookies("ProductIDList")("ProductID"),0)
		
		If Len(strProductID) = 0 Then
			Exit Function
		Else
			SQL = "SELECT shopid,NowPrice FROM [ECCMS_ShopList] WHERE ChannelID=" & ChannelID & " And isAccept > 0 And shopid in (" & strProductID & ")"
			Set Rs = enchiasp.Execute(SQL)
			
			If Rs.BOF And Rs.EOF Then
				Set Rs = Nothing
				Exit Function
			Else
				TotalPrice = 0
				Do While Not Rs.EOF
					QuantityID = "Quantity_" & Rs("shopid")
					Quantity = enchiasp.ChkNumeric(Request.Cookies("ProductIDList")(QuantityID))
					If Quantity = 0 Then Quantity = 1
					
					UnitPrice = Rs("NowPrice") * Quantity
					TotalPrice = TotalPrice + UnitPrice
					Rs.MoveNext
				Loop
			End If
			Set Rs = Nothing
		End If
		
		CountTotalMoney = CCur(TotalPrice )
	End Function
	'=================================================
	'��������CheckEmail
	'��  �ã��ж�EMAIL
	'=================================================
	Public Function CheckEmail(Byval email)
		Dim names, ename, i, c
		CheckEmail = True
		email = Trim(email)
		names = Split(email, "@")
		If UBound(names) <> 1 Then
			CheckEmail = False
			Exit Function
		End If
		For Each ename in names
			If Len(ename) <= 0 Then
				CheckEmail = False
				Exit Function
			End If
			For i = 1 To Len(ename)
				c = LCase(Mid(ename, i, 1))
				If InStr("abcdefghijklmnopqrstuvwxyz_-.", c) <= 0 And Not IsNumeric(c) Then
					CheckEmail = False
					Exit Function
				End If
			Next
			If Left(ename, 1) = "." Or Right(ename, 1) = "." Then
				CheckEmail = False
				Exit Function
			End If
		Next
		If InStr(names(1), ".") <= 0 Then
			CheckEmail = False
			Exit Function
		End If
		i = Len(names(1)) - InStrRev(names(1), ".")
		If i <> 2 And i <> 3 Then
			CheckEmail = False
			Exit Function
		End If
		If InStr(email, "..") > 0 Then
			CheckEmail = False
		End If
	End Function
	'///---�����ύ���̽���
	'-------------------------------------------------
	'///---����֧�����ع��̿�ʼ
	'=================================================
	'��������BuildReceive
	'��  �ã�����֧������ҳ��
	'=================================================
	Public Sub BuildReceive()
		Dim strContent,errcode
		Dim wp,arrChinaeBank,ErrorMsg
		Dim OrderForm,PaymentMoney
		Dim ServiceCharge,BuyMoney
		Dim Consignee,Readme
		Dim userid,UserName

		On Error Resume Next
		
		userid = Clng(enchiasp.memberid)
		UserName = enchiasp.CheckRequest(enchiasp.membername,45)
		skinid = CLng(enchiasp.ChannelSkin)
		enchiasp.LoadTemplates ChannelID, 6, skinid

		'--����Ȩ������
		If CInt(enchiasp.GroupSetting(30)) = 0 Then
			Call OutAlertScript(enchiasp.CheckStr(enchiasp.HtmlSetting(8)))
			Exit Sub
		End If

		HtmlContent = enchiasp.HtmlContent
		HtmlContent = Replace(HtmlContent, "{$ChannelRootDir}", ChannelRootDir)
		HtmlContent = Replace(HtmlContent, "{$InstallDir}", strInstallDir)
		HtmlContent = Replace(HtmlContent, "{$ChannelID}", ChannelID)
		HtmlContent = Replace(HtmlContent, "{$ModuleName}", enchiasp.ModuleName)
		HtmlContent = Replace(HtmlContent, "{$ShopIndex}", strIndexName)
		
		arrChinaeBank = Split(enchiasp.ChinaeBank, "|||")

		Set wp = New WebPayment_Cls
		wp.PayPlatform = CInt(enchiasp.StopBankPay)		'--ѡ������֧������
		wp.Paymentid = Trim(arrChinaeBank(0))			'--����֧��ID
		wp.Paymentkey = Trim(arrChinaeBank(1))			'--����֧��KEY
		wp.Percent = enchiasp.CheckNumeric(arrChinaeBank(2))	'--����֧��������
		wp.ReceivePage						'--ִ������֧��
		OrderForm = enchiasp.CheckInfuse(wp.Orderid,35)		'--���ض�����
		PaymentMoney = CCur(wp.Paymoney)			'--����֧�����
		BuyMoney = CCur(wp.Buymoney)				'--���ؼ�ȥ�����Ѻ�Ľ��
		ServiceCharge = CCur(wp.ServiceCharge)			'--����������
		Consignee = enchiasp.CheckInfuse(wp.Consignee,35)		'--���ض���������
		Readme = enchiasp.CheckRequest(wp.Comment,220)		'--���ض���˵��
		errcode = CInt(wp.ErrNumber)				'--���ش�����룬0=�ɹ�
		If errcode > 0 Then ErrorMsg = wp.Description		'--���ش�����Ϣ
		'--���鷵�ض������Ƿ���ȷ
		Set Rs = enchiasp.Execute("SELECT id,totalmoney,PayDone FROM ECCMS_OrderForm WHERE OrderID='"& OrderForm &"'")
		If Rs.BOF And Rs.EOF Then
			ErrorMsg ="�Ƿ������������Ų���ȷ������"
			errcode = 1
		Else
			'--������صĽ����ύʱ�Ľ��������ش���
			If BuyMoney <> Rs("totalmoney") Then
				ErrorMsg ="�Ƿ����������׽��ԣ�����"
				errcode = 1
			End If
			'--�ж��Ƿ��ظ��ύ����
			If Rs("PayDone") > 0 Then
				ErrorMsg ="�˴ν����Ѿ��ɹ�,�벻Ҫ�ظ��ύ���ݣ�����"
				errcode = 1
			End If
		End If
		Set Rs = Nothing
		If errcode = 0 Then
			'--�����֧���ɹ�
			HtmlContent = Replace(HtmlContent, "{$PageTitle}", "����֧���ɹ�")
			strContent = enchiasp.HtmlSetting(13)
			strContent = Replace(strContent, "{$ReceiveTiele}", "�ɹ�")
			strContent = Replace(strContent, "{$OrderID}", OrderForm)
			strContent = Replace(strContent, "{$PayMoney}", FormatNumber(PaymentMoney,2,-1))
			strContent = Replace(strContent, "{$BuyMoney}", FormatNumber(BuyMoney,2,-1))
			strContent = Replace(strContent, "{$ServiceCharge}", FormatNumber(ServiceCharge,2,-1))
			strContent = Replace(strContent, "{$DateTime}", Now())
			strContent = Replace(strContent, "{$ErrorMsg}", "")
			strContent = Replace(strContent, "{$PayState}", "��ϲ������������ɡ�")
			'--֧���ɹ���ʼ�������ݿ�state
			'--���¶�������״̬
			enchiasp.Execute ("UPDATE ECCMS_OrderForm SET Paymode='����֧��',PayDone=1 WHERE OrderID='"& OrderForm &"'")
			'--����ǻ�Ա���»�Ա���Ѽ�¼
			If enchiasp.memberid > 0 Then
				enchiasp.Execute ("UPDATE ECCMS_User SET prepaid=prepaid+" & BuyMoney & " WHERE UserName='"& UserName &"' And userid=" & userid)
			End If
			'--�����ؽ�����ϸ��
			If Trim(Consignee) = "" Then
				If userid > 0 Then
					Consignee = Username
				Else
					Consignee = "�����û�"
				End If
			End If
			Set Rs = CreateObject("ADODB.Recordset")
			SQL = "SELECT * FROM ECCMS_Account WHERE (AccountID is null)"
			Rs.Open SQL,Conn,1,3
			Rs.addnew
				Rs("payer").Value = Consignee
				Rs("payee").Value = enchiasp.CheckRequest(enchiasp.SiteName,20)
				Rs("product").Value = "���Ϲ���"
				Rs("Amount").Value = 1
				Rs("unit").Value = "��"
				Rs("price").Value = BuyMoney
				Rs("TotalPrices").Value = PaymentMoney
				Rs("DateAndTime").Value = Now()
				Rs("Accountype").Value = 0
				Rs("Explain").Value = Readme
				Rs("Reclaim").Value = 0
			Rs.update
			Rs.Close:Set Rs = Nothing
		Else
			'--���֧��ʧ��
			HtmlContent = Replace(HtmlContent, "{$PageTitle}", "����֧��ʧ��")
			strContent = enchiasp.HtmlSetting(16)
			strContent = Replace(strContent, "{$ErrorMsg}", ErrorMsg)
			strContent = Replace(strContent, "{$ReceiveTiele}", "ʧ��")
			strContent = Replace(strContent, "{$OrderID}", OrderForm)
			strContent = Replace(strContent, "{$PayMoney}", FormatNumber(PaymentMoney,2,-1))
			strContent = Replace(strContent, "{$BuyMoney}", FormatNumber(BuyMoney,2,-1))
			strContent = Replace(strContent, "{$ServiceCharge}", FormatNumber(ServiceCharge,2,-1))
			strContent = Replace(strContent, "{$DateTime}", Now())
			strContent = Replace(strContent, "{$PayState}", "�Բ��𣡱�����ʧ�ܡ�")
		End If
		Set wp = Nothing
		
		HtmlContent = Replace(HtmlContent, "{$PublicContent}", strContent)
		ReplaceString
		Response.Write HtmlContent
	End Sub
	'///---����֧�����ع��̽���
	'=================================================
	'��������ShowPaginate
	'��  �ã��̳��ղؼз�ҳ
	'=================================================
	Public Function ShowPaginate(ByVal str,ByVal CurrentPage,ByVal Pcount,ByVal totalrec,ByVal maxperpage)
		Dim strTempage
		strTempage = str
		strTempage = Replace(strTempage, "{$CurrentPage}", CurrentPage)
		strTempage = Replace(strTempage, "{$PageCount}", Pcount)
		strTempage = Replace(strTempage, "{$Totalrec}", totalrec)
		strTempage = Replace(strTempage, "{$totalrec}", totalrec)
		strTempage = Replace(strTempage, "{$MaxPerPage}", maxperpage)
		ShowPaginate = strTempage
	End Function
	'///---��Ʒ�ղؼй��̿�ʼ
	'=================================================
	'��������BuildFavorite
	'��  �ã��̳��ղؼ�
	'=================================================
	Public Sub BuildFavorite()
		
		Dim strContent,ErrorMsg
		Dim userid,UserName,Action,i,j
		Dim maxfavsize,strTopNum,strPagination
		Dim maxperpage,CurrentPage,Pcount,totalrec,page_count
		Dim StyleList,favcount,FavoriteHint,shopid
		
		userid = Clng(enchiasp.memberid)
		UserName = enchiasp.CheckRequest(enchiasp.membername,45)
		Action = LCase(enchiasp.CheckInfuse(Request("action"),8))
		skinid = CLng(enchiasp.ChannelSkin)
		maxfavsize = CLng(enchiasp.GroupSetting(36))

		If maxfavsize > 0 Then
			strTopNum = "TOP " & maxfavsize
			FavoriteHint = "�����ղؼ������Դ�� <font color=""red""><b>" & maxfavsize & "</b></font> ����Ʒ��"
		Else
			strTopNum = ""
			FavoriteHint = "�����ղؼд�С�����ƣ�"
		End If
		'--Ȩ������
		If userid = 0 Or UserName = "" Then
			Call OutAlertScript("�Բ���ֻ��ע���Ա����ʹ���ղع��ܡ�")
			Exit Sub
		End If
		If CInt(enchiasp.GroupSetting(35)) = 0 Then
			Call OutAlertScript("�Բ�����û��ʹ���ղؼе�Ȩ�ޡ�")
			Exit Sub
		End If
		
		Select Case Action
		Case "del"
			If userid = 0 Or enchiasp.ChkNumeric(Request("favid")) = 0 Then
				ErrorMsg = "��û��ѡ���ղ�ID��������û�е�¼��"
				FoundErr = True
			End If
			If FoundErr = False Then
				enchiasp.Execute("DELETE FROM ECCMS_Favourite WHERE userid="& userid &" And favid="& enchiasp.ChkNumeric(Request("favid")))
				Response.Redirect("favorite.asp")
			End If
		Case "add"
			If FoundErr = False Then
				Call AddFavorite
				Response.Redirect("favorite.asp")
			End If
		Case "modify"
			shopid = enchiasp.ChkNumeric(Request("shopid"))
			If userid = 0 Or shopid = 0 Then
				ErrorMsg = "��û��ѡ����ƷID��������û�е�¼��"
				FoundErr = True
			End If
			If FoundErr = False Then
				SQL = "SELECT shopid,TradeName,NowPrice FROM [ECCMS_ShopList] WHERE ChannelID=" & ChannelID & " And isAccept > 0 And shopid=" & shopid
				Set Rs = enchiasp.Execute(SQL)
				If Not (Rs.BOF And Rs.EOF) Then
					enchiasp.Execute ("UPDATE ECCMS_Favourite SET ProductName='"& enchiasp.CheckStr(Rs("TradeName")) &"',UnitPrice="& Rs("NowPrice") &" WHERE userid="& userid &" And shopid="& Rs("shopid"))
				End If
				Set Rs = Nothing
				Response.Redirect("favorite.asp")
			End If
		End Select
		'--ͳ���ղؼд�С
		Set Rs = enchiasp.Execute("SELECT COUNT(favid) FROM ECCMS_Favourite WHERE userid="& userid)
		favcount = CLng(Rs(0))
		'�����ж�Ϊ�Զ�ɾ�����������Ʒ
		If favcount > maxfavsize And maxfavsize <> 0 Then
			i = favcount - maxfavsize
			SQL = "SELECT TOP "& i &" favid FROM ECCMS_Favourite WHERE userid="& userid &" ORDER BY favid DESC"
			Set Rs=enchiasp.Execute(SQL)
			While Not Rs.EOF
				enchiasp.Execute("DELETE FROM ECCMS_Favourite WHERE favid="& rs(0))
				Rs.movenext
			Wend
			smsCount = Maxsms
		End if
		Rs.Close:Set Rs = Nothing
		
		enchiasp.LoadTemplates ChannelID, 6, skinid
		
		maxperpage = CInt(enchiasp.HtmlSetting(20))	'--ÿҳ��
		CurrentPage = enchiasp.ChkNumeric(Request("page"))
		If CurrentPage = 0 Then CurrentPage = 1
			
		HtmlContent = enchiasp.HtmlContent
		HtmlContent = Replace(HtmlContent, "{$ChannelRootDir}", ChannelRootDir)
		HtmlContent = Replace(HtmlContent, "{$InstallDir}", strInstallDir)
		HtmlContent = Replace(HtmlContent, "{$ChannelID}", ChannelID)
		HtmlContent = Replace(HtmlContent, "{$ModuleName}", enchiasp.ModuleName)
		HtmlContent = Replace(HtmlContent, "{$ShopIndex}", strIndexName)
		HtmlContent = Replace(HtmlContent, "{$PageTitle}", "�ղؼ�")
		
		If FoundErr = False Then
			strContent = enchiasp.HtmlSetting(17)
			Set Rs = CreateObject("ADODB.Recordset")
			SQL = "SELECT " & strTopNum & " favid,userid,shopid,ProductName,UnitPrice,addTime FROM [ECCMS_Favourite] WHERE userid="& userid &" ORDER BY favid DESC"
			Rs.Open SQL, Conn, 1, 1
			If Rs.BOF And Rs.EOF Then
				strContent = strContent & enchiasp.HtmlSetting(22)
			Else
				Rs.PageSize = maxperpage
				Rs.AbsolutePage = CurrentPage
				page_count = 0
				totalrec = Rs.recordcount
				j = (CurrentPage - 1) * maxperpage + 1
				Do While Not Rs.EOF And (Not page_count = Rs.PageSize)
					If (page_count Mod 2) = 0 Then
						StyleList = 1
					Else
						StyleList = 2
					End If
					strContent = strContent & enchiasp.HtmlSetting(18)
					strContent = Replace(strContent, "{$FavoriteID}", Rs("favid"))
					strContent = Replace(strContent, "{$ProductID}", Rs("shopid"))
					strContent = Replace(strContent, "{$TradeName}", Rs("ProductName"))
					strContent = Replace(strContent, "{$FavouriteID}", Rs("favid"))
					strContent = Replace(strContent, "{$UnitPrice}", FormatNumber(Rs("UnitPrice"),2,-1))
					strContent = Replace(strContent, "{$AddTime}", Rs("addTime"))
					strContent = Replace(strContent, "{$StyleList}", StyleList)
					strContent = Replace(strContent, "{$Ordered}", j)
					Rs.movenext
					page_count = page_count + 1
					j = j + 1
					If page_count >= maxperpage Then Exit Do
				Loop
			End if
			Rs.Close:Set Rs = Nothing
			'--��ҳ����
			If totalrec Mod maxperpage = 0 Then
				Pcount =  totalrec \ maxperpage
			Else
				Pcount =  totalrec \ maxperpage + 1
			End If
			If page_count = 0 Then CurrentPage = 0
			'--��ʾ�������
			strPagination = enchiasp.HtmlSetting(21)
			strPagination = ShowPaginate(strPagination,CurrentPage,Pcount,totalrec,maxperpage)
			'--��ҳ�������
			strContent = strContent & enchiasp.HtmlSetting(19)
			strContent = Replace(strContent, "{$Pagination}", strPagination)
		Else
			strContent = enchiasp.HtmlSetting(14)
			strContent = Replace(strContent, "{$ErrorMsg}", ErrorMsg)
			strContent = Replace(strContent, "{$DateTime}", Now())
			
		End If
		strContent = Replace(strContent, "{$MaxFavourite}", maxfavsize)
		strContent = Replace(strContent, "{$FavoriteHint}", FavoriteHint)
		HtmlContent = Replace(HtmlContent, "{$PublicContent}", strContent)
		ReplaceString
		Response.Write HtmlContent
	End Sub
	'=================================================
	'��������AddFavorite
	'��  �ã�����ղؼ�
	'=================================================
	Private Sub AddFavorite()
		Dim ValueFavourite,strProductID
		Dim favcount,shopid,userid
		
		userid = Clng(enchiasp.memberid)
		shopid = enchiasp.ChkNumeric(Request("shopid"))
		
		If shopid = 0 Then
			strProductID = enchiasp.CheckRequest(Request.Cookies("ProductIDList")("ProductID"),0)
			If Len(strProductID) = 0 Then Exit Sub 

			SQL = "SELECT shopid,TradeName,NowPrice FROM [ECCMS_ShopList] WHERE ChannelID=" & ChannelID & " And isAccept > 0 And shopid in (" & strProductID & ")"
			Set Rs = enchiasp.Execute(SQL)
			If Not (Rs.BOF And Rs.EOF) Then
				Do While Not Rs.EOF
					favcount = enchiasp.Execute("SELECT COUNT(favid) FROM ECCMS_Favourite WHERE userid="& userid &" And shopid=" & Rs("shopid"))(0)
					If CLng(favcount) = 0 Then
						ValueFavourite = "" & userid & "," & Rs("shopid") & ",'" & enchiasp.CheckStr(Rs("TradeName")) & "'," & Rs("NowPrice") & "," & NowString & ""
						SQL = "INSERT INTO ECCMS_Favourite (userid,shopid,ProductName,UnitPrice,addTime) values ("& ValueFavourite &")"
						enchiasp.Execute(SQL)
					End if
					Rs.movenext
				Loop
			End if
			Rs.Close:Set Rs = Nothing
		Else
			If shopid = 0 Then Exit Sub
			SQL = "SELECT shopid,TradeName,NowPrice FROM [ECCMS_ShopList] WHERE ChannelID=" & ChannelID & " And isAccept > 0 And shopid=" & shopid
			Set Rs = enchiasp.Execute(SQL)
			If Not (Rs.BOF And Rs.EOF) Then
					
				favcount = enchiasp.Execute("SELECT COUNT(favid) FROM ECCMS_Favourite WHERE userid="& userid &" And shopid=" & Rs("shopid"))(0)
				If CLng(favcount) = 0 Then
					ValueFavourite = "" & userid & "," & Rs("shopid") & ",'" & enchiasp.CheckStr(Rs("TradeName")) & "'," & Rs("NowPrice") & "," & NowString & ""
					SQL = "INSERT INTO ECCMS_Favourite (userid,shopid,ProductName,UnitPrice,addTime) values ("& ValueFavourite &")"
					enchiasp.Execute(SQL)
				End if
			End if
			Rs.Close:Set Rs = Nothing
		End If
	End Sub
	'///---�̳��ղؼй��̽���
	'-------------------------------------------------
	'///---������ѯ���̿�ʼ
	'=================================================
	'��������BuildOrderQuery
	'��  �ã�������ѯ
	'=================================================
	Public Sub BuildOrderQuery()
		Dim strContent,ErrorMsg
		Dim TotalPrice,keyword
		
		keyword = enchiasp.CheckInfuse(Request("word"),35)
		'blurry = enchiasp.ChkNumeric(Request("blurry"))
		skinid = CLng(enchiasp.ChannelSkin)
		enchiasp.LoadTemplates ChannelID, 6, skinid
		
		HtmlContent = enchiasp.HtmlContent
		HtmlContent = Replace(HtmlContent, "{$ChannelRootDir}", ChannelRootDir)
		HtmlContent = Replace(HtmlContent, "{$InstallDir}", strInstallDir)
		HtmlContent = Replace(HtmlContent, "{$ChannelID}", ChannelID)
		HtmlContent = Replace(HtmlContent, "{$ModuleName}", enchiasp.ModuleName)
		HtmlContent = Replace(HtmlContent, "{$ShopIndex}", strIndexName)
		HtmlContent = Replace(HtmlContent, "{$PageTitle}", "������ѯ")
		
		If Trim(keyword) <> "" Then
			Set Rs = enchiasp.Execute("SELECT id,OrderID,Surcharge,totalmoney,Consignee,PayMode,addTime,finish,PayDone FROM ECCMS_OrderForm WHERE OrderID='"& Trim(keyword) &"'")
			If Rs.BOF And Rs.EOF Then
				strContent = enchiasp.HtmlSetting(24)
				strContent = Replace(strContent, "{$QueryInfo}", enchiasp.HtmlSetting(25))
			Else
				strContent = enchiasp.HtmlSetting(23)
				TotalPrice = FormatNumber(Rs("totalmoney") + Rs("Surcharge"))
				strContent = Replace(strContent, "{$TotalPrice}", TotalPrice)
				strContent = Replace(strContent, "{$OrderID}", Rs("OrderID"))
				strContent = Replace(strContent, "{$Surcharge}", FormatNumber(Rs("Surcharge")))
				strContent = Replace(strContent, "{$totalmoney}", FormatNumber(Rs("totalmoney")))
				strContent = Replace(strContent, "{$Consignee}", Rs("Consignee"))
				strContent = Replace(strContent, "{$AddTime}", Rs("addTime"))
				strContent = Replace(strContent, "{$PayMode}", Rs("PayMode"))
				If Rs("finish") > 0 Then
					strContent = Replace(strContent, "{$OrderState}", "<font color=""blue"">�Ѵ���</font>")
				Else
					strContent = Replace(strContent, "{$OrderState}", "<font color=""red"">δ����</font>")
				End If
				If Rs("PayDone") > 0 Then
					strContent = Replace(strContent, "{$PayState}", "<font color=""blue"">��֧��</font>")
				Else
					strContent = Replace(strContent, "{$PayState}", "<font color=""red"">δ֧��</font>")
				End If
			End If
			Rs.Close:Set Rs = Nothing
		Else
			strContent = enchiasp.HtmlSetting(24)
			strContent = Replace(strContent, "{$QueryInfo}", "")
		End If

		HtmlContent = Replace(HtmlContent, "{$PublicContent}", strContent)
		ReplaceString
		Response.Write HtmlContent
	End Sub
	'///---������ѯ���̽���
	'///----------------------------------------------
	'///---�̳ǰ������̿�ʼ
	'=================================================
	'��������BuildHelpInfo
	'��  �ã��̳ǰ�����Ϣ
	'=================================================
	Public Sub BuildHelpInfo()
		Dim strContent,HelpContent
		
		On Error Resume Next
		
		skinid = CLng(enchiasp.ChannelSkin)
		enchiasp.LoadTemplates ChannelID, 6, skinid
		
		HtmlContent = enchiasp.HtmlContent
		HtmlContent = Replace(HtmlContent, "{$ChannelRootDir}", ChannelRootDir)
		HtmlContent = Replace(HtmlContent, "{$InstallDir}", strInstallDir)
		HtmlContent = Replace(HtmlContent, "{$ChannelID}", ChannelID)
		HtmlContent = Replace(HtmlContent, "{$ModuleName}", enchiasp.ModuleName)
		HtmlContent = Replace(HtmlContent, "{$ShopIndex}", strIndexName)
		HtmlContent = Replace(HtmlContent, "{$IndexTitle}", strIndexName)
		HtmlContent = Replace(HtmlContent, "{$PageTitle}", "��������")
		
		strContent = enchiasp.HtmlSetting(26)
		If InStr(strContent,"{$HelpContent}") > 0 Then
			HelpContent = enchiasp.Readfile("help.inc")
			strContent = Replace(strContent, "{$HelpContent}", HelpContent)
		End if
		HtmlContent = Replace(HtmlContent, "{$PublicContent}", strContent)
		ReplaceString
		Response.Write HtmlContent
	End Sub
	'///---�̳������б�ʼ,��:������Ʒ,�Ƽ���Ʒ,������Ʒ
	'-- ������Ʒ�б�
	Public Sub BuildNewProductList()
		CurrentPage = enchiasp.ChkNumeric(Request("page"))
		If CurrentPage = 0 Then CurrentPage = 1
		Response.Write LoadOtherList(3)
	End Sub
	'-- ������Ʒ�б�
	Public Sub BuildHotProductList()
		CurrentPage = enchiasp.ChkNumeric(Request("page"))
		If CurrentPage = 0 Then CurrentPage = 1
		Response.Write LoadOtherList(2)
	End Sub
	'-- �Ƽ���Ʒ�б�
	Public Sub BuildBestProductList()
		CurrentPage = enchiasp.ChkNumeric(Request("page"))
		If CurrentPage = 0 Then CurrentPage = 1
		Response.Write LoadOtherList(1)
	End Sub
	'=================================================
	'��������LoadOtherShopList
	'��  �ã������������̳��б�
	'=================================================
	Public Function LoadOtherList(t)
		On Error Resume Next
		Dim HtmlFileName, SQL1, SQL2

		skinid = CLng(enchiasp.ChannelSkin)
		enchiasp.LoadTemplates ChannelID, 5, skinid
		HtmlFilePath = enchiasp.InstallDir & enchiasp.ChannelDir & "special/"
		HtmlContent = enchiasp.HtmlContent
		HtmlContent = Replace(HtmlContent, "{$ChannelRootDir}", ChannelRootDir)
		HtmlContent = Replace(HtmlContent, "{$InstallDir}", strInstallDir)
		HtmlContent = Replace(HtmlContent, "{$ChannelID}", ChannelID)
		HtmlContent = Replace(HtmlContent, "{$ModuleName}", enchiasp.ModuleName)
		HtmlContent = Replace(HtmlContent, "{$ShopIndex}", strIndexName)
		PageType = 3
		If CInt(t) = 1 Then
			strClassName = enchiasp.HtmlSetting(10)
			SQL1 = "And IsBest>0"
			SQL2 = "And A.IsBest>0 ORDER BY A.addTime DESC,A.shopid DESC"
		ElseIf CInt(t) = 2 Then
			strClassName = enchiasp.HtmlSetting(11)			
			SQL1 = ""
			SQL2 = "ORDER BY A.AllHits DESC,A.addTime DESC,A.shopid DESC"
		Else
			strClassName = enchiasp.HtmlSetting(12)
			SQL1 = ""
			SQL2 = "ORDER BY A.addTime DESC ,A.shopid DESC"
		End If
		HtmlContent = Replace(HtmlContent, "{$PageTitle}", strClassName)
		Call ReplaceString
		maxperpage = CLng(enchiasp.HtmlSetting(1))
		If CLng(CurrentPage) = 0 Then CurrentPage = 1
		'��¼����
		TotalNumber = enchiasp.Execute("SELECT COUNT(shopid) FROM ECCMS_ShopList WHERE ChannelID = " & ChannelID & " And isAccept>0  " & SQL1 & "")(0)
		If TotalNumber >= CLng(enchiasp.HtmlSetting(5)) Then TotalNumber = CLng(enchiasp.HtmlSetting(5))
		TotalPageNum = CLng(TotalNumber / maxperpage)  '�õ���ҳ��
		If TotalPageNum < TotalNumber / maxperpage Then TotalPageNum = TotalPageNum + 1
		If CurrentPage < 1 Then CurrentPage = 1
		If CurrentPage > TotalPageNum Then CurrentPage = TotalPageNum
		Set Rs = CreateObject("ADODB.Recordset")
		SQL = "SELECT TOP " & CLng(enchiasp.HtmlSetting(5)) & " A.ShopID,A.ClassID,A.TradeName,A.Explain,A.PastPrice,A.NowPrice,A.star,ProductImage,A.addTime,A.AllHits,A.HtmlFileDate,A.isBest,C.ClassName,C.ParentID,C.ParentStr,C.skinid,C.HtmlFileDir,C.ChildStr,C.UseHtml FROM [ECCMS_ShopList] A INNER JOIN [ECCMS_Classify] C On A.ClassID=C.ClassID where A.ChannelID=" & ChannelID & " And A.isAccept>0 " & SQL2
		If isSqlDataBase = 1 Then
			Set Rs = enchiasp.Execute(SQL)
		Else
			Rs.Open SQL, Conn, 1, 1
		End If

		If Rs.BOF And Rs.EOF Then
			HtmlContent = Replace(HtmlContent, "{$ReadListPage}", "��û���ҵ��κ�" & enchiasp.ModuleName & "")
			HtmlContent = Replace(HtmlContent, enchiasp.CutFixContent(HtmlContent, "[ShowRepetend]", "[/ShowRepetend]", 1), "")
			If CreateHtml <> 0 Then
				enchiasp.CreatPathEx (HtmlFilePath)
				HtmlFileName = HtmlFilePath & ReadListPageName(ClassID, CurrentPage)
				enchiasp.CreatedTextFile HtmlFileName, HtmlContent
				If IsShowFlush = 1 Then 
					Response.Write "<li style=""font-size: 12px;"">����" & enchiasp.ModuleName & "�б�HTML���... <a href=" & HtmlFileName & " target=_blank>" & Server.MapPath(HtmlFileName) & "</a></li>" & vbNewLine
					Response.Flush
				End If
			End If
		Else
			TempListContent = enchiasp.CutFixContent(HtmlContent, "[ShowRepetend]", "[/ShowRepetend]", 1)
			If CreateHtml <> 0 Then
				LoadOtherListHtml(n)
			Else
				Call LoadShopAspList
			End If
		End If
		Rs.Close: Set Rs = Nothing
		HtmlContent = HTML.ReadFriendLink(HtmlContent)
		'LoadOtherList = HtmlContent
		If CreateHtml = 0 Then LoadOtherList = HtmlContent
	End Function
	'================================================
	'��������LoadOtherListHtml
	'��  �ã�װ�������б�����HTML
	'================================================
	Private Sub LoadOtherListHtml(t)
		Dim HtmlFileName, sulCurrentPage
		Dim Perownum,ii,w
		
		If IsNull(TempListContent) Then Exit Sub
		On Error Resume Next

		Perownum = enchiasp.ChkNumeric(enchiasp.HtmlSetting(4))

		enchiasp.CreatPathEx (HtmlFilePath)
		For CurrentPage = n To TotalPageNum
			Rs.MoveFirst
			i = 0
			If TotalPageNum > 1 Then Rs.Move (CurrentPage - 1) * maxperpage
			ListContent = ""
			j = (CurrentPage - 1) * maxperpage + 1
			If Perownum > 1 Then 
				ListContent = enchiasp.HtmlSetting(6)
				w = FormatPercent(100 / Perownum / 100,0)
			End If
			
			Do While Not Rs.EOF And i < CInt(maxperpage)
				If Not Response.IsClientConnected Then Response.end
				
				If Perownum > 1 Then
					ListContent = ListContent & "<tr valign=""top"">" & vbCrLf
					For ii = 1 To Perownum
						ListContent = ListContent & "<td width=""" & w & """class=""shoplistrow"">"
						If Not Rs.EOF Then
							Call LoadListDetail
							Rs.movenext
							i = i + 1
							j = j + 1
						End If
						ListContent = ListContent & "</td>" & vbCrLf
					Next
					ListContent = ListContent & "</tr>" & vbCrLf
				Else
					Call LoadListDetail
					Rs.MoveNext
					i = i + 1
					j = j + 1
				End If
				
				If i >= maxperpage Then Exit Do
			Loop
			
			Dim strHtmlFront, strHtmlPage
			If t = 1 Then
				sulCurrentPage = enchiasp.HtmlPrefix & "Best"
			ElseIf t = 2 Then
				sulCurrentPage = enchiasp.HtmlPrefix & "Hot"
			Else
				sulCurrentPage = enchiasp.HtmlPrefix & "New"
			End If
			strHtmlFront = sulCurrentPage
			strHtmlPage = ShowHtmlPage(CurrentPage, TotalPageNum, TotalNumber, maxperpage, strHtmlFront, enchiasp.HtmlExtName, strClassName)
			HtmlTemplate = HtmlContent
			HtmlTemplate = Replace(HtmlTemplate, TempListContent, ListContent)
			HtmlTemplate = Replace(HtmlTemplate, "{$ReadListPage}", strHtmlPage)
			HtmlTemplate = Replace(HtmlTemplate, "[ShowRepetend]", "")
			HtmlTemplate = Replace(HtmlTemplate, "[/ShowRepetend]", "")
			'��ʼ�����ӷ����HTMLҳ
			HtmlFileName = HtmlFilePath & sulCurrentPage & enchiasp.Supplemental(CurrentPage, 3) & enchiasp.HtmlExtName
			enchiasp.CreatedTextFile HtmlFileName, HtmlTemplate
			If IsShowFlush = 1 Then 
				Response.Write "<li style=""font-size: 12px;"">����" & enchiasp.ModuleName & "�б�HTML���... <a href=" & HtmlFileName & " target=_blank>" & Server.MapPath(HtmlFileName) & "</a></li>" & vbNewLine
				Response.Flush
			End If
		Next

	End Sub
	'================================================
	'��������LoadOtherListDetail
	'��  �ã�װ��������Ʒ�б�ϸ��
	'================================================
	Private Sub LoadOtherListDetail()
		Dim sTitle, sTopic, TradeName, ListStyle
		Dim ShopUrl, ShopTime, sClassName
		Dim ProductImageUrl, ProductImage,ProductIntro

		ListContent = ListContent & TempListContent
		If (i Mod 2) = 0 Then
			ListStyle = 1
		Else
			ListStyle = 2
		End If
		
		sTitle = Rs("TradeName")
		On Error Resume Next
		If CInt(CreateHtml) <> 0 Then
			ShopUrl = ChannelRootDir & Rs("HtmlFileDir") & enchiasp.ShowDatePath(Rs("HtmlFileDate"), enchiasp.HtmlPath) & ReadPagination(0)
			sClassName = ChannelRootDir & Rs("HtmlFileDir")
		Else
			ShopUrl = ChannelRootDir & "show.asp?id=" & Rs("shopid")
			sClassName = ChannelRootDir & "list.asp?classid=" & Rs("ClassID")
		End If
		ProductImageUrl = enchiasp.GetImageUrl(Rs("ProductImage"), enchiasp.ChannelDir)
		ProductImage = enchiasp.GetFlashAndPic(ProductImageUrl, CInt(enchiasp.HtmlSetting(4)), CInt(enchiasp.HtmlSetting(5)))
		ProductImage = "<a href='" & ShopUrl & "' title='" & Rs("TradeName") & "'>" & ProductImage & "</a>"
		sClassName = "<a href='" & sClassName & "' title='" & Rs("ClassName") & "'>" & Rs("ClassName") & "</a>"
		TradeName = "<a href='" & ShopUrl & "' title='" & Rs("TradeName") & "' class=showtopic>" & sTitle & "</a>"

		ProductIntro = enchiasp.CutString(Rs("Explain"), CInt(enchiasp.HtmlSetting(3)))
		ProductIntro = enchiasp.JAPEncode(ProductIntro)
		
		ShopTime = enchiasp.ShowDateTime(Rs("addTime"), CInt(enchiasp.HtmlSetting(2)))
		ListContent = Replace(ListContent, "{$ClassifyName}", sClassName)
		ListContent = Replace(ListContent, "{$TradeName}", TradeName)
		ListContent = Replace(ListContent, "{$ShopTopic}", sTitle)
		ListContent = Replace(ListContent, "{$ShopUrl}", ShopUrl)
		ListContent = Replace(ListContent, "{$ProductImage}", ProductImage)
		ListContent = Replace(ListContent, "{$ShopID}", Rs("shopid"))
		ListContent = Replace(ListContent, "{$ShopHits}", Rs("AllHits"))
		ListContent = Replace(ListContent, "{$Star}", Rs("star"))
		ListContent = Replace(ListContent, "{$ShopDateTime}", ShopTime)
		ListContent = Replace(ListContent, "{$ProductIntro}", ProductIntro)
		ListContent = Replace(ListContent, "{$ListStyle}", ListStyle)
		ListContent = Replace(ListContent, "{$PastPrice}", FormatNumber(Rs("PastPrice"),2,-1))
		ListContent = Replace(ListContent, "{$NowPrice}", FormatNumber(Rs("NowPrice"),2,-1))
		ListContent = Replace(ListContent, "{$IsTop}", Rs("IsTop"))
		ListContent = Replace(ListContent, "{$IsBest}", Rs("IsBest"))
		ListContent = Replace(ListContent, "{$Order}", j)
	End Sub
	'///---�̳������б����
	'///---�̳ǰ������̽���
	'#############################\\ִ�������б�ʼ//#############################
	'=================================================
	'��������BuildProductSearch
	'��  �ã���ʾ�̳�����ҳ��
	'=================================================
	Public Sub BuildProductSearch()
		Dim SearchMaxPageList
		Dim Action, findword,keyword
		Dim rsClass, strNoResult, s

		PageType = 5
		keyword = enchiasp.ChkQueryStr(Trim(Request("keyword")))
		keyword = enchiasp.CheckInfuse(keyword,200)
		s = enchiasp.ChkNumeric(Request.QueryString("s"))
		
		If keyword = "" Then
			Call OutAlertScript("������Ҫ��ѯ�Ĺؼ��֣�")
			Exit Sub
		End If
		
		If Not enchiasp.CheckQuery(keyword) Then
			Call OutAlertScript("���ѯ�Ĺؼ����зǷ��ַ���\n�뷵����������ؼ��ֲ�ѯ��")
			Exit Sub
		End If

		skinid = CLng(enchiasp.ChannelSkin)
		
		On Error Resume Next
		
		enchiasp.LoadTemplates ChannelID, 7, skinid
		
		If enchiasp.HtmlSetting(4) <> "0" Then
			If IsNumeric(enchiasp.HtmlSetting(4)) Then
				SearchMaxPageList = CLng(enchiasp.HtmlSetting(4))
			Else
				SearchMaxPageList = 50
			End If
		Else
			SearchMaxPageList = 50
		End If

		If enchiasp.strLength(keyword) < CLng(enchiasp.HtmlSetting(5)) Or enchiasp.strLength(keyword) > CLng(enchiasp.HtmlSetting(6)) Then
			Call OutAlertScript("��ѯ����\n����ѯ�Ĺؼ��ֲ���С�� " & enchiasp.HtmlSetting(5) & " ���ߴ��� " & enchiasp.HtmlSetting(6) & " ���ֽڡ�")
			Exit Sub
		End If

		strNoResult = Replace(enchiasp.HtmlSetting(8), "{$KeyWord}", keyword)
		Action = enchiasp.CheckStr(Trim(Request("act")))
		Action = enchiasp.CheckInfuse(Action)

		
		If LCase(Action) = "topic" Then
			findword = "A.TradeName like '%" & keyword & "%'"
		ElseIf LCase(Action) = "content" Then
			If CInt(enchiasp.FullContQuery) <> 0 Then
				findword = "A.Explain like '%" & keyword & "%'"
			Else
				Call OutAlertScript(Replace(Replace(enchiasp.HtmlSetting(10), Chr(34), "\"""), vbCrLf, ""))
				Exit Sub
			End If
		Else
			findword = "A.TradeName like '%" & keyword & "%'"
		End If

		If IsEmpty(Session("QueryLimited")) Then
			Session("QueryLimited") = keyword & "|" & Action & "|" & Now()
		Else
			Dim QueryLimited
			QueryLimited = Split(Session("QueryLimited"), "|")
			If UBound(QueryLimited) = 2 Then
				If CStr(Trim(QueryLimited(0))) = CStr(keyword) And CStr(Trim(QueryLimited(1))) = CStr(Action) Then
					Session("QueryLimited") = keyword & "|" & Action & "|" & Now()
				Else
					If DateDiff("s", QueryLimited(2), Now()) < CLng(enchiasp.HtmlSetting(7)) Then
						Dim strLimited
						strLimited = Replace(enchiasp.HtmlSetting(9), "{$TimeLimited}", enchiasp.HtmlSetting(7))
						Call OutAlertScript(Replace(Replace(strLimited, Chr(34), "\"""), vbCrLf, ""))
						Exit Sub
					Else
						Session("QueryLimited") = keyword & "|" & Action & "|" & Now()
					End If
				End If
			Else
				Session("QueryLimited") = keyword & "|" & Action & "|" & Now()
			End If
		End If

		HtmlContent = enchiasp.HtmlContent
		HtmlContent = Replace(HtmlContent, "{$ChannelRootDir}", ChannelRootDir)
		HtmlContent = Replace(HtmlContent, "{$InstallDir}", strInstallDir)
		HtmlContent = Replace(HtmlContent, "{$ChannelID}", ChannelID)
		HtmlContent = Replace(HtmlContent, "{$ModuleName}", enchiasp.ModuleName)
		HtmlContent = Replace(HtmlContent, "{$ShopIndex}", strIndexName)
		HtmlContent = Replace(HtmlContent, "{$KeyWord}", KeyWord)
		HtmlContent = Replace(HtmlContent, "{$PageTitle}", enchiasp.ModuleName & "����")
		HtmlContent = Replace(HtmlContent, "{$QueryKeyWord}", "<font color=red><strong>" & keyword & "</strong></font>")
		Call ReplaceString
		If IsNumeric(Request("classid")) And Request("classid") <> "" Then
			Set rsClass = enchiasp.Execute("SELECT ClassID,ChildStr FROM [ECCMS_Classify] WHERE ChannelID = " & ChannelID & " And ClassID=" & CLng(Request("classid")))
			If rsClass.BOF And rsClass.EOF Then
				HtmlContent = Replace(HtmlContent, "{$ReadListPage}", strNoResult, 1, 1, 1)
				HtmlContent = Replace(HtmlContent, "{$ReadListPage}", "")
				HtmlContent = Replace(HtmlContent, "{$totalrec}", 0)
				HtmlContent = Replace(HtmlContent, enchiasp.CutFixContent(HtmlContent, "[ShowRepetend]", "[/ShowRepetend]", 1), "")
				Set rsClass = Nothing
				Response.Write HtmlContent
				Exit Sub
			Else
				findword = "A.ClassID IN (" & rsClass("ChildStr") & ") And " & findword
			End If
			rsClass.Close: Set rsClass = Nothing
		End If
		maxperpage = CInt(enchiasp.HtmlSetting(1))
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

		Set Rs = CreateObject("ADODB.Recordset")
		SQL = "SELECT TOP " & SearchMaxPageList & " A.ShopID,A.ClassID,A.TradeName,A.Explain,A.PastPrice,A.NowPrice,A.star,ProductImage,A.addTime,A.AllHits,A.HtmlFileDate,C.ClassName,C.HtmlFileDir,C.UseHtml FROM [ECCMS_ShopList] A INNER JOIN [ECCMS_Classify] C On A.ClassID=C.ClassID where A.ChannelID=" & ChannelID & " And A.isAccept > 0 And " & findword & " ORDER BY A.addTime DESC ,A.shopid DESC"
		Rs.Open SQL, Conn, 1, 1
		If Rs.BOF And Rs.EOF Then
			'���û���ҵ��������,��������õı�ǩ����
			HtmlContent = Replace(HtmlContent, "{$ReadListPage}", strNoResult, 1, 1, 1)
			HtmlContent = Replace(HtmlContent, "{$ReadListPage}", "")
			HtmlContent = Replace(HtmlContent, "{$totalrec}", 0)
			HtmlContent = Replace(HtmlContent, enchiasp.CutFixContent(HtmlContent, "[ShowRepetend]", "[/ShowRepetend]", 1), "")
		Else
			TotalNumber = Rs.RecordCount
			If (TotalNumber Mod maxperpage) = 0 Then
				TotalPageNum = TotalNumber \ maxperpage
			Else
				TotalPageNum = TotalNumber \ maxperpage + 1
			End If
			If CurrentPage > TotalPageNum Then CurrentPage = TotalPageNum
			If CurrentPage < 1 Then CurrentPage = 1
			HtmlContent = Replace(HtmlContent, "{$totalrec}", TotalNumber)
			'��ȡģ���ǩ[ShowRepetend][/ReadSoftList]�е��ַ���
			TempListContent = enchiasp.CutFixContent(HtmlContent, "[ShowRepetend]", "[/ShowRepetend]", 1)
			Call LoadSearchList
		End If
		Rs.Close: Set Rs = Nothing
		Response.Write HtmlContent
		Exit Sub
	End Sub
	'================================================
	'��������LoadSearchList
	'��  �ã�װ�������б�
	'================================================

	Private Sub LoadSearchList()
		If IsNull(TempListContent) Then Exit Sub
		i = 0
		If TotalPageNum > 1 Then Rs.Move (CurrentPage - 1) * maxperpage
		j = (CurrentPage - 1) * maxperpage + 1
		ListContent = ""
		Do While Not Rs.EOF And i < CInt(maxperpage)
			If Not Response.IsClientConnected Then Response.End
			Call SearchResult
			Rs.MoveNext
			i = i + 1
			j = j + 1
			If i >= maxperpage Then Exit Do
		Loop
		Dim strPagination
		strPagination = ShowListPage(CurrentPage, TotalPageNum, TotalNumber, maxperpage, ASPCurrentPage(PageType), "�������")
		HtmlContent = Replace(HtmlContent, TempListContent, ListContent)
		HtmlContent = Replace(HtmlContent, "[ShowRepetend]", "")
		HtmlContent = Replace(HtmlContent, "[/ShowRepetend]", "")
		HtmlContent = Replace(HtmlContent, "{$ReadListPage}", strPagination)
	End Sub
	'================================================
	'��������SearchResult
	'��  �ã�װ�������б���ϸ
	'================================================
	Private Sub SearchResult()
		Dim sTitle, sTopic, TradeName, ListStyle
		Dim ShopUrl, ShopTime, sClassName
		Dim ProductImageUrl, ProductImage,ProductIntro

		ListContent = ListContent & TempListContent
		If (i Mod 2) = 0 Then
			ListStyle = 1
		Else
			ListStyle = 2
		End If
		
		sTitle = Rs("TradeName")
		On Error Resume Next
		If CInt(CreateHtml) <> 0 Then
			ShopUrl = ChannelRootDir & Rs("HtmlFileDir") & enchiasp.ShowDatePath(Rs("HtmlFileDate"), enchiasp.HtmlPath) & ReadPagination(0)
			sClassName = ChannelRootDir & Rs("HtmlFileDir")
		Else
			ShopUrl = ChannelRootDir & "show.asp?id=" & Rs("shopid")
			sClassName = ChannelRootDir & "list.asp?classid=" & Rs("ClassID")
		End If
		ProductImageUrl = enchiasp.GetImageUrl(Rs("ProductImage"), enchiasp.ChannelDir)
		ProductImage = enchiasp.GetFlashAndPic(ProductImageUrl, CInt(enchiasp.HtmlSetting(11)), CInt(enchiasp.HtmlSetting(12)))
		ProductImage = "<a href='" & ShopUrl & "' title='" & Rs("TradeName") & "'>" & ProductImage & "</a>"
		sClassName = "<a href='" & sClassName & "' title='" & Rs("ClassName") & "'>" & Rs("ClassName") & "</a>"
		TradeName = "<a href='" & ShopUrl & "' title='" & Rs("TradeName") & "' class=showtopic>" & sTitle & "</a>"

		ProductIntro = enchiasp.CutString(Rs("Explain"), CInt(enchiasp.HtmlSetting(3)))
		
		ShopTime = enchiasp.ShowDateTime(Rs("addTime"), CInt(enchiasp.HtmlSetting(2)))
		ListContent = Replace(ListContent, "{$ClassifyName}", sClassName)
		ListContent = Replace(ListContent, "{$totalrec}", TotalNumber)
		ListContent = Replace(ListContent, "{$TradeName}", TradeName)
		ListContent = Replace(ListContent, "{$ShopTopic}", sTitle)
		ListContent = Replace(ListContent, "{$ShopUrl}", ShopUrl)
		ListContent = Replace(ListContent, "{$ProductImage}", ProductImage)
		ListContent = Replace(ListContent, "{$ShopID}", Rs("shopid"))
		ListContent = Replace(ListContent, "{$ShopHits}", Rs("AllHits"))
		ListContent = Replace(ListContent, "{$Star}", Rs("star"))
		ListContent = Replace(ListContent, "{$ShopDateTime}", ShopTime)
		ListContent = Replace(ListContent, "{$ProductIntro}", ProductIntro)
		ListContent = Replace(ListContent, "{$ListStyle}", ListStyle)
		ListContent = Replace(ListContent, "{$PastPrice}", FormatNumber(Rs("PastPrice"),2,-1))
		ListContent = Replace(ListContent, "{$NowPrice}", FormatNumber(Rs("NowPrice"),2,-1))
		ListContent = Replace(ListContent, "{$IsTop}", Rs("IsTop"))
		ListContent = Replace(ListContent, "{$IsBest}", Rs("IsBest"))
		ListContent = Replace(ListContent, "{$Order}", j)
	End Sub
	'================================================
	'��������ProductComment
	'��  �ã���Ʒ����
	'================================================
	Private Function ProductComment(shopid)
		Dim rsComment, SQL, strContent, strComment
		Dim i, Resize, strRearrange
		Dim ArrayTemp()

		On Error Resume Next
		Set rsComment = enchiasp.Execute("SELECT TOP " & CInt(enchiasp.HtmlSetting(5)) & " content,Grade,username,postime,postip FROM ECCMS_Comment WHERE ChannelID=" & ChannelID & " And postid = " & shopid & " ORDER BY postime DESC,CommentID DESC")
		If Not (rsComment.EOF And rsComment.BOF) Then
			i = 0
			Resize = 0
			Do While Not rsComment.EOF
				ReDim Preserve ArrayTemp(i + Resize)
				strContent = ArrayTemp(i) & enchiasp.HtmlSetting(7)
				strComment = enchiasp.CutString(rsComment("content"), CInt(enchiasp.HtmlSetting(6)))
				strContent = Replace(strContent, "{$Comment}", enchiasp.HTMLEncode(strComment))
				strContent = Replace(strContent, "{$UserName}", enchiasp.HTMLEncode(rsComment("username")))
				strContent = Replace(strContent, "{$UserGrade}", rsComment("Grade"))
				strContent = Replace(strContent, "{$postime}", rsComment("postime"))
				strContent = Replace(strContent, "{$postip}", rsComment("postip"))
				ArrayTemp(i) = strContent
				rsComment.MoveNext
				i = i + 1
			Loop
		End If
		rsComment.Close
		strRearrange = Join(ArrayTemp, vbCrLf)
		Set rsComment = Nothing
		ProductComment = strRearrange
	End Function
	'================================================
	'��������ReplaceString
	'��  �ã��滻ģ������
	'================================================
	Private Sub ReplaceString()
		HtmlContent = Replace(HtmlContent, "{$SelectedType}", "")
		HtmlContent = ReadClassMenu(HtmlContent)
		HtmlContent = ReadClassMenubar(HtmlContent)
		HtmlContent = HTML.ReadShopPic(HtmlContent)
		HtmlContent = HTML.ReadShopList(HtmlContent)
		HtmlContent = Replace(HtmlContent, "{$ChannelRootDir}", ChannelRootDir)
		HtmlContent = Replace(HtmlContent,"{$ShopIndex}", strIndexName)
		HtmlContent = Replace(HtmlContent, "{$SkinPath}", enchiasp.SkinPath)
		HtmlContent = Replace(HtmlContent, "{$InstallDir}", strInstallDir)
		HtmlContent = Replace(HtmlContent, "{$ChannelID}", ChannelID)
	End Sub
	'================================================
	'��������BuildShopComment
	'��  �ã���ʾ��Ʒ����
	'================================================
	Public Sub BuildShopComment()
		Dim TradeName, HtmlFileUrl, HtmlFileName
		Dim AverageGrade, TotalComment, TempListContent
		Dim strComment, strCheckBox, strAdminComment

		enchiasp.PreventInfuse
		strCheckBox = ""
		strAdminComment = ""
		On Error Resume Next
		shopid = enchiasp.ChkNumeric(Request("shopid"))
		If shopid = 0 Then
			Response.Write "<Br><Br><Br>Sorry�������ϵͳ����,��ѡ����ȷ�����ӷ�ʽ��"
			Response.End
		End If
		skinid = CLng(enchiasp.ChannelSkin)
		enchiasp.LoadTemplates ChannelID, 8, skinid
		HtmlContent = enchiasp.HtmlContent
		HtmlContent = Replace(HtmlContent, "{$ChannelRootDir}", ChannelRootDir)
		HtmlContent = Replace(HtmlContent, "{$InstallDir}", strInstallDir)
		HtmlContent = Replace(HtmlContent, "{$ChannelID}", ChannelID)
		HtmlContent = Replace(HtmlContent, "{$ModuleName}", enchiasp.ModuleName)
		HtmlContent = Replace(HtmlContent, "{$ShopIndex}", strIndexName)
		HtmlContent = Replace(HtmlContent, "{$IndexTitle}", strIndexName)
		HtmlContent = Replace(HtmlContent, "{$PageTitle}", enchiasp.ModuleName & "����")
		HtmlContent = Replace(HtmlContent, "{$shopid}", shopid)
		HtmlContent = Replace(HtmlContent, "{$ShopID}", shopid)
		HtmlContent = Replace(HtmlContent, "{$UserName}", enchiasp.membername)
		HtmlContent = Replace(HtmlContent, "{$UserName}", "")


		'��ñ���
		SQL = "SELECT TOP 1 A.shopid,A.ClassID,A.TradeName,A.HtmlFileDate,A.ForbidEssay,C.HtmlFileDir,C.UseHtml FROM [ECCMS_ShopList] A INNER JOIN [ECCMS_Classify] C ON A.ClassID=C.ClassID WHERE A.ChannelID=" & ChannelID & " And A.isAccept > 0 And A.shopid = " & shopid
		Set Rs = enchiasp.Execute(SQL)
		If Rs.EOF And Rs.BOF Then
			Response.Write "�Ѿ�û����"
			Set Rs = Nothing
			Exit Sub
		Else
			If CreateHtml <> 0 Then
				HtmlFileUrl = ChannelRootDir & Rs("HtmlFileDir") & enchiasp.ShowDatePath(Rs("HtmlFileDate"), enchiasp.HtmlPath)
				HtmlFileName = enchiasp.ReadFileName(Rs("HtmlFileDate"), Rs("shopid"), enchiasp.HtmlExtName, enchiasp.HtmlPrefix, enchiasp.HtmlForm, "")
				TradeName = "<a href=" & HtmlFileUrl & HtmlFileName & ">" & Rs("TradeName") & "</a>"
			Else
				TradeName = "<a href=show.asp?id=" & Rs("shopid") & ">" & Rs("TradeName") & "</a>"
			End If
			ForbidEssay = Rs("ForbidEssay")
		End If
		Rs.Close
		Set Rs = CreateObject("adodb.recordset")
		SQL = "SELECT COUNT(CommentID) As TotalComment,AVG(Grade) As avgGrade FROM ECCMS_Comment WHERE ChannelID=" & ChannelID & " And postid = " & shopid
		Set Rs = enchiasp.Execute(SQL)
		TotalComment = Rs("TotalComment")
		AverageGrade = Round(Rs("avgGrade"))
		If IsNull(AverageGrade) Then AverageGrade = 0
		Rs.Close: Set Rs = Nothing
		HtmlContent = Replace(HtmlContent, "{$TradeName}", TradeName)
		HtmlContent = Replace(HtmlContent, "{$TotalComment}", TotalComment)
		HtmlContent = Replace(HtmlContent, "{$AverageGrade}", AverageGrade)
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
		'ÿҳ��ʾ������
		maxperpage = CInt(enchiasp.PaginalNum)
		'��¼����
		TotalNumber = TotalComment
		TotalPageNum = CInt(TotalNumber / maxperpage)  '�õ���ҳ��
		If TotalPageNum < TotalNumber / maxperpage Then TotalPageNum = TotalPageNum + 1
		If CurrentPage < 1 Then CurrentPage = 1
		If CurrentPage > TotalPageNum Then CurrentPage = TotalPageNum
		Set Rs = CreateObject("ADODB.Recordset")
		SQL = "SELECT * FROM ECCMS_Comment WHERE ChannelID=" & ChannelID & " And postid = " & shopid & " ORDER BY postime DESC,CommentID DESC"
		If isSqlDataBase = 1 Then
			Set Rs = enchiasp.Execute(SQL)
		Else
			Rs.Open SQL, Conn, 1, 1
		End If
		If Rs.BOF And Rs.EOF Then
			'���û���ҵ��������,��������õı�ǩ����
			HtmlContent = Replace(HtmlContent, "{$ReadListPage}", "��ʱ���˲μ�����", 1, 1, 1)
			HtmlContent = Replace(HtmlContent, "{$ReadListPage}", "")
			HtmlContent = Replace(HtmlContent, enchiasp.CutFixContent(HtmlContent, "[ShowRepetend]", "[/ShowRepetend]", 1), "")
		Else
			Rs.MoveFirst
			i = 0
			If TotalPageNum > 1 Then Rs.Move (CurrentPage - 1) * maxperpage
			ListContent = ""
			'��ȡģ���ǩ[ShowRepetend][/ReadArticleList]�е��ַ���
			TempListContent = enchiasp.CutFixContent(HtmlContent, "[ShowRepetend]", "[/ShowRepetend]", 1)
			Do While Not Rs.EOF And i < CInt(maxperpage)
				If Not Response.IsClientConnected Then Response.End
				ListContent = ListContent & TempListContent
				strComment = enchiasp.HTMLEncode(Rs("Content"))
				ListContent = Replace(ListContent, "{$CommentContent}", strComment)
				ListContent = Replace(ListContent, "{$UserName}", enchiasp.HTMLEncode(Rs("username")))
				ListContent = Replace(ListContent, "{$CommentGrade}", Rs("Grade"))
				ListContent = Replace(ListContent, "{$PostTime}", Rs("postime"))
				ListContent = Replace(ListContent, "{$PostIP}", Rs("postip"))
				If Session("AdminName") <> "" Or enchiasp.membergrade = "999" Then
					strCheckBox = "<input type='checkbox' name='selCommentID' value='" & Rs("CommentID") & "'>"
				End If
				ListContent = Replace(ListContent, "{$SelCheckBox}", strCheckBox)
				Rs.MoveNext
				i = i + 1
				If i >= maxperpage Then Exit Do
			Loop
		End If
		Rs.Close: Set Rs = Nothing
		HtmlContent = Replace(HtmlContent, TempListContent, ListContent)
		HtmlContent = Replace(HtmlContent, "[ShowRepetend]", "")
		HtmlContent = Replace(HtmlContent, "[/ShowRepetend]", "")
		If Session("AdminName") <> "" Or enchiasp.membergrade = "999" Then
			strAdminComment = "<input class=Button type=button name=chkall value='ȫѡ' onClick=""CheckAll(this.form)""><input class=Button type=button name=chksel value='��ѡ' onClick=""ContraSel(this.form)"">" & vbNewLine
			strAdminComment = strAdminComment & "<input type=hidden name=shopid value='" & shopid & "'>" & vbNewLine
			strAdminComment = strAdminComment & "<input type=hidden name=action value='del'>" & vbNewLine
			strAdminComment = strAdminComment & "<input class=Button type=submit name=Submit2 value='ɾ��ѡ�е�����' onclick=""{if(confirm('��ȷ��ִ�иò�����?')){this.document.selform.submit();return true;}return false;}"">"
		End If
		HtmlContent = Replace(HtmlContent, "{$AdminComment}", strAdminComment)
		Call ShowCommentPage
		Call ReplaceString
		If enchiasp.CheckStr(LCase(Request.Form("action"))) = "del" Then
			Call CommentDel
		End If
		If enchiasp.CheckStr(LCase(Request.Form("action"))) = "save" Then
			Call CommentSave
		End If
		Response.Write HtmlContent
		
	End Sub
	'================================================
	'��������ShowCommentPage
	'��  �ã����۷�ҳ
	'================================================
	Private Sub ShowCommentPage()
		Dim FileName, ii, n, strTemp

		FileName = "comment.asp"
		On Error Resume Next
		If TotalNumber Mod maxperpage = 0 Then
			n = TotalNumber \ maxperpage
		Else
			n = TotalNumber \ maxperpage + 1
		End If
		strTemp = "<table cellspacing=1 width='100%' border=0><tr><td align=center> " & vbCrLf
		If CurrentPage < 2 Then
			strTemp = strTemp & " �������� <font COLOR=#FF0000>" & TotalNumber & "</font> ��&nbsp;&nbsp;�� ҳ&nbsp;&nbsp;��һҳ&nbsp;&nbsp;&nbsp;"
		Else
			strTemp = strTemp & "�������� <font COLOR=#FF0000>" & TotalNumber & "</font> ��&nbsp;&nbsp;<a href=" & FileName & "?page=1&shopid=" & Request("shopid") & ">�� ҳ</a>&nbsp;&nbsp;"
			strTemp = strTemp & "<a href=" & FileName & "?page=" & CurrentPage - 1 & "&shopid=" & Request("shopid") & ">��һҳ</a>&nbsp;&nbsp;&nbsp;"
		End If

		If n - CurrentPage < 1 Then
			strTemp = strTemp & "��һҳ&nbsp;&nbsp;β ҳ " & vbCrLf
		Else
			strTemp = strTemp & "<a href=" & FileName & "?page=" & (CurrentPage + 1) & "&shopid=" & Request("shopid") & ">��һҳ</a>"
			strTemp = strTemp & "&nbsp;&nbsp;<a href=" & FileName & "?page=" & n & "&shopid=" & Request("shopid") & ">β ҳ</a>" & vbCrLf
		End If
		strTemp = strTemp & "&nbsp;ҳ�Σ�<strong><font color=red>" & CurrentPage & "</font>/" & n & "</strong>ҳ "
		strTemp = strTemp & "&nbsp;<b>" & maxperpage & "</b>��/ҳ " & vbCrLf
		strTemp = strTemp & "</td></tr></table>" & vbCrLf
		HtmlContent = Replace(HtmlContent, "{$ReadListPage}", strTemp)
	End Sub
	'================================================
	'��������CommentDel
	'��  �ã�����ɾ��
	'================================================
	Private Sub CommentDel()
		Dim selCommentID

		If enchiasp.CheckPost = False Then
			Call OutAlertScript("���ύ�����ݲ��Ϸ����벻Ҫ���ⲿ�ύ����")
			Exit Sub
		End If
		If Not IsEmpty(Request.Form("selCommentID")) Then
			selCommentID = enchiasp.CheckStr(Request("selCommentID"))
			If Session("AdminName") <> "" Or enchiasp.membergrade = "999" Then
				enchiasp.Execute ("delete from ECCMS_Comment where ChannelID=" & ChannelID & " And CommentID in (" & selCommentID & ")")
				Call OutHintScript("����ɾ���ɹ���")
			Else
				Call OutAlertScript("�Ƿ���������û��ɾ�����۵�Ȩ�ޡ�")
				Exit Sub
			End If
		End If
	End Sub
	'================================================
	'��������CommentSave
	'��  �ã�������ӱ���
	'================================================
	Public Sub CommentSave()
		If enchiasp.CheckPost = False Then
			Call OutAlertScript("���ύ�����ݲ��Ϸ����벻Ҫ���ⲿ�ύ����")
			Exit Sub
		End If
		On Error Resume Next
		Call PreventRefresh
		If CInt(enchiasp.AppearGrade) <> 0 And Session("AdminName") = "" Then
			If CInt(enchiasp.AppearGrade) > CInt(enchiasp.membergrade) Then
				Call OutAlertScript("��û�з������۵�Ȩ�ޣ�������ǻ�Ա���½���ٲ������ۡ�")
				Exit Sub
			End If
		End If
		If ForbidEssay <> 0 Then
			Call OutAlertScript("��" & enchiasp.ModuleName & "��ֹ�������ۣ�")
			Exit Sub
		End If
		If Trim(Request.Form("UserName")) = "" Then
			Call OutAlertScript("�û�������Ϊ�գ�")
			Exit Sub
		End If
		If Len(Trim(Request.Form("UserName"))) > 15 Then
			Call OutAlertScript("�û������ܴ���15���ַ���")
			Exit Sub
		End If
		If enchiasp.strLength(Request.Form("content")) < enchiasp.LeastString Then
			Call OutAlertScript("�������ݲ���С��" & enchiasp.LeastString & "�ַ���")
			Exit Sub
		End If
		If enchiasp.strLength(Request.Form("content")) > enchiasp.MaxString Then
			Call OutAlertScript("�������ݲ��ܴ���" & enchiasp.MaxString & "�ַ���")
			Exit Sub
		End If
		shopid = enchiasp.ChkNumeric(Request.Form("shopid"))
		Set Rs = CreateObject("ADODB.RecordSet")
		SQL = "SELECT * FROM ECCMS_Comment WHERE (CommentID is null)"
		Rs.Open SQL, Conn, 1, 3
		Rs.AddNew
			Rs("ChannelID") = ChannelID
			Rs("postid") = shopid
			Rs("UserName") = Trim(Request.Form("UserName"))
			Rs("Grade") = Trim(Request.Form("Grade"))
			Rs("content") = Request.Form("content")
			Rs("postime") = Now()
			Rs("postip") = enchiasp.GetUserip
		Rs.Update
		Rs.Close: Set Rs = Nothing
		If CreateHtml <> 0 Then LoadShopInfo(shopid)
		Session("UserRefreshTime") = Now()
		Response.Redirect (Request.ServerVariables("HTTP_REFERER"))
		Exit Sub
	End Sub
	Public Sub PreventRefresh()
		Dim RefreshTime

		RefreshTime = 20
		If DateDiff("s", Session("UserRefreshTime"), Now()) < RefreshTime Then
			Response.Write "<META http-equiv=Content-Type content=text/html; chaRset=gb2312><meta HTTP-EQUIV=REFRESH CONTENT=" & RefreshTime & "><br>��ҳ�������˷�ˢ�»��ƣ��벻Ҫ��" & RefreshTime & "��������ˢ�±�ҳ��<BR>���ڴ�ҳ�棬���Ժ󡭡�"
			Response.End
		End If
	End Sub
	Private Function ReadPagination(n)
		Dim HtmlFileName, CurrentPage
		
		CurrentPage = n
		HtmlFileName = enchiasp.ReadFileName(Rs("HtmlFileDate"), Rs("shopid"), enchiasp.HtmlExtName, enchiasp.HtmlPrefix, enchiasp.HtmlForm, CurrentPage)
		ReadPagination = HtmlFileName
	End Function
	Private Function ReadListPageName(ClassID, CurrentPage)
		ReadListPageName = enchiasp.ClassFileName(ClassID, enchiasp.HtmlExtName, enchiasp.HtmlPrefix, CurrentPage)
	End Function
	Public Function ASPCurrentPage(stype)
		Dim CurrentUrl
		Select Case stype
			Case "1"
				CurrentUrl = "&amp;classid=" & Trim(Request("classid"))
			Case "2"
				CurrentUrl = "&amp;sid=" & Trim(Request("sid"))
			Case "3"
				CurrentUrl = ""
			Case "4"
				CurrentUrl = ""
			Case "6"
				CurrentUrl = "&amp;type=" & enchiasp.CheckStr(Request("type"))
			Case Else
				If Trim(Request("word")) <> "" Then
					CurrentUrl = "&amp;word=" & Trim(Request("word"))
				Else
					CurrentUrl = "&amp;act=" & Trim(Request("act")) & "&amp;classid=" & Trim(Request("classid")) & "&amp;keyword=" & Trim(Request("keyword"))
				End If
		End Select
		ASPCurrentPage = CurrentUrl
	End Function
End Class

%>
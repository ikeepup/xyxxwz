
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
Class WebPayment_Cls
	Private sPaymentid, sPaymentkey, sReturnurl, sPlatform, sPstate
	Private sOrderid, sPaymoney, sMoneytype, sLanguage
	Private sComment, sRemark, sConsigner, submit_value
	Private sConsignee, sAddress, sPostcode, sTelephone, sEmail
	Public ErrNumber, mPercent
	Private strChinaeBank, sEncrypt
	Private PaymentContent

	Private Sub Class_Initialize()
		On Error Resume Next
		sPaymentid = "1051011239"
		sPaymentkey = "enchicom778899"
		sReturnurl = "http://www.enchi.com.cn/user/Receive.asp"
		sPlatform = 1
		sOrderid = "2005428-1301-5188"
		sPaymoney = "100.00"
		sMoneytype = 1
		sLanguage = 1
		sPstate = 0
		sComment = "在线支付"
		sRemark = "在线支付"
		sConsigner = "Consigner"
		sConsignee = "Consignee"
		sAddress = "运城"
		sPostcode = "51800"
		sTelephone = "0359-8698845"
		sEmail = "liuyunfan@163.com"
		submit_value = "进入在线支付平台"
		sEncrypt = "12345678"
		mPercent = 0
	End Sub

	Private Sub Class_Terminate()
		
	End Sub
	'---- 错误代码
	Public Property Get Description()
		Select Case ErrNumber
			Case 1: Description = "定单号错误。"
			Case 2: Description = "付款金额错误！"
			Case 3: Description = "认证签名不能为空值!"
			Case 4: Description = "认证信息出错，此次交易失败！！！"
			Case 5: Description = "认证信息出错，请不要重复提交数据，此次交易失败！！！"
			Case 6: Description = "对不起！本站暂未开通在线支付功能,请选择其它方法支付"
			Case 7: Description = "错误的系统参数"
			Case 8: Description = "本站暂未开通在线支付功能,或者本站程序没有注册,请选择其它方法支付"
			Case Else
				Description = Empty
		End Select
	End Property
	'---- 手续费百分比
	Public Property Let Percent(ByVal NewValue)
		mPercent = NewValue
	End Property
	'---- 支付平台
	Public Property Let PayPlatform(ByVal NewValue)
		sPlatform = NewValue
	End Property
	'---- 支付ID
	Public Property Let Paymentid(ByVal NewValue)
		sPaymentid = NewValue
		If Trim(sPaymentid) = "" Then
			sPaymentid = "1051011239"
		End If
	End Property
	Public Property Get Paymentid()
		Paymentid = sPaymentid
	End Property
	'---- 支付KEY
	Public Property Let Paymentkey(ByVal NewValue)
		sPaymentkey = NewValue
		If Trim(sPaymentkey) = "" Then
			sPaymentkey = "enchiasp778899"
		End If
	End Property
	'---- 返回URL
	Public Property Let Returnurl(ByVal NewValue)
		sReturnurl = NewValue
	End Property
	'---- 交易状态
	Public Property Let Pstate(ByVal NewValue)
		sPstate = NewValue
	End Property
	Public Property Get Pstate()
		Pstate = sPstate
	End Property
	'--- 定单号
	Public Property Let Orderid(ByVal NewValue)
		sOrderid = NewValue
	End Property
	Public Property Get Orderid()
		Orderid = sOrderid
	End Property
	'---- 支付金额
	Public Property Let Paymoney(ByVal NewValue)
		sPaymoney = ReadPayMoney(NewValue, False)
	End Property
	Public Property Get Paymoney()
		Paymoney = sPaymoney
	End Property
	'---- 交易金额
	Public Property Get Buymoney()
		Buymoney = ReadPayMoney(sPaymoney, True)
	End Property
	'---- 手续费
	Public Property Get ServiceCharge()
		ServiceCharge = sPaymoney - ReadPayMoney(sPaymoney, True)
	End Property
	'---- 支付币种
	Public Property Let Moneytype(ByVal NewValue)
		sMoneytype = NewValue
	End Property
	Public Property Get Moneytype()
		Moneytype = sMoneytype
	End Property
	'---- 支付语言
	Public Property Let Planguage(ByVal NewValue)
		sLanguage = NewValue
	End Property
	'---- 支付备注
	Public Property Let Comment(ByVal NewValue)
		sComment = NewValue
	End Property
	Public Property Get Comment()
		Comment = sComment
	End Property
	'---- 支付备注
	Public Property Let Remark(ByVal NewValue)
		sRemark = NewValue
	End Property
	Public Property Get Remark()
		Remark = sRemark
	End Property
	'---- 收货人名称
	Public Property Let Consignee(ByVal NewValue)
		sConsignee = NewValue
	End Property
	Public Property Get Consignee()
		Consignee = sConsignee
	End Property
	'---- 收货人地址
	Public Property Let Address(ByVal NewValue)
		sAddress = NewValue
	End Property
	Public Property Get Address()
		Address = sAddress
	End Property
	'---- 收货人邮编
	Public Property Let Postcode(ByVal NewValue)
		sPostcode = NewValue
	End Property
	Public Property Get Postcode()
		Postcode = sPostcode
	End Property
	'---- 收货人电话
	Public Property Let Telephone(ByVal NewValue)
		sTelephone = NewValue
	End Property
	Public Property Get Telephone()
		Telephone = sTelephone
	End Property
	'---- 收货人E_Mail
	Public Property Let Email(ByVal NewValue)
		sEmail = NewValue
	End Property
	Public Property Get Email()
		Email = sEmail
	End Property
	'---- 发货人
	Public Property Let Consigner(ByVal NewValue)
		sConsigner = NewValue
	End Property
	Public Property Get Consigner()
		Consigner = sConsigner
	End Property
	'---- 提交按钮
	Public Property Let submitvalue(ByVal NewValue)
		submit_value = NewValue
	End Property
	'---- 加密密码
	Public Property Let Encrypt(ByVal NewValue)
		sEncrypt = NewValue
	End Property
	'================================================
	'过程名：GetWebSiteUrl
	'作  用：取得带端口的URL
	'================================================
	Public Property Get GetWebSiteUrl()
		If Request.ServerVariables("SERVER_PORT") = "80" Then
			GetWebSiteUrl = "http://" & Request.ServerVariables("server_name")
		Else
			GetWebSiteUrl = "http://" & Request.ServerVariables("server_name") & ":" & Request.ServerVariables("SERVER_PORT")
		End If
	End Property
	'================================================
	'过程名：PaymentPlatform
	'作  用：在线支付平台
	'================================================
	Public Sub PaymentPlatform()
		On Error Resume Next
		If sPlatform = 1 Then
			Call Payment_nps
		ElseIf sPlatform = 2 Then
			Call Payment_chinabank
		Else
			ErrNumber = 6
			Exit Sub
		End If
		Response.Write PaymentContent
	End Sub
	'================================================
	'函数名：ShowPayment
	'作  用：显示在线支付平台
	'================================================
	Public Function ShowPayment()
		On Error Resume Next
		If sPlatform = 1 Then
			Call Payment_nps
		ElseIf sPlatform = 2 Then
			Call Payment_chinabank
		Else
			ShowPayment = ""
			ErrNumber = 6
			Exit Function
		End If
		ShowPayment = PaymentContent
	End Function
	'================================================
	'过程名：payment_nps
	'作  用：NPS在线支付平台
	'================================================
	Private Sub Payment_nps()
		On Error Resume Next
		Dim digest, OrderMessage
		Dim m_url, m_orderid, m_oamount, modate, m_ocomment
		Dim m_ocurrency, m_language, s_postcode, s_tel, s_eml, r_postcode, r_tel, r_eml
		m_orderid = Trim(sOrderid)                      '---- 定单号
		m_oamount = sPaymoney                           '---- 金 额
		m_url = sReturnurl                              '---- 返回URL
		m_ocurrency = sMoneytype                        '---- 币    种
		m_language = sLanguage                          '---- 语言选择
		s_postcode = sPostcode                          '---- 消费者邮码
		s_tel = sTelephone                              '---- 消费者邮码
		s_eml = sEmail                                  '---- 消费者邮件
		r_postcode = sPostcode                          '---- 收货人电话
		r_tel = sTelephone                              '---- 收货人电话
		r_eml = sEmail                                  '---- 收货人邮件
		m_ocomment = sComment                           '---- 备 注
		modate = Date                                   '---- 日 期
		
		OrderMessage = sPaymentid & m_orderid & m_oamount & m_ocurrency & m_url & m_language & s_postcode & s_tel & s_eml & r_postcode & r_tel & r_eml & modate & sPaymentkey
		digest = UCase(Trim(md5(OrderMessage,True)))
		
		PaymentContent = "<table>         <tr>" & vbNewLine
		PaymentContent = PaymentContent & "<form method=""post"" action=""https://payment.nps.cn/VirReceiveMerchantAction.do"" name=""payform"" target=""_blank"">" & vbNewLine
		PaymentContent = PaymentContent & "        <td>" & vbNewLine
		PaymentContent = PaymentContent & "<input type=""hidden"" Name=""M_ID"" value=""" & sPaymentid & """>" & vbNewLine                         '---- 商 家 号
		PaymentContent = PaymentContent & "<input type=""hidden"" Name=""MOrderID"" value=""" & m_orderid & """>" & vbNewLine                       '---- 订 单 号
		PaymentContent = PaymentContent & "<input type=""hidden"" Name=""MOAmount"" value=""" & m_oamount & """>" & vbNewLine                       '---- 订单金额
		PaymentContent = PaymentContent & "<input type=""hidden"" Name=""MOCurrency"" value=""" & m_ocurrency & """>" & vbNewLine                   '---- 币    种
		PaymentContent = PaymentContent & "<input type=""hidden"" name=""M_URL"" value=""" & m_url & """>" & vbNewLine                              '---- 返回地址
		PaymentContent = PaymentContent & "<input type=""hidden"" Name=""M_Language"" value=""" & m_language & """>" & vbNewLine                    '---- 语言选择
		PaymentContent = PaymentContent & "<input type=""hidden"" Name=""S_Name"" value=""" & sConsignee & """>" & vbNewLine                        '---- 消费者姓名
		PaymentContent = PaymentContent & "<input type=""hidden"" Name=""S_Address"" value=""" & sAddress & """>" & vbNewLine                       '---- 消费者住址
		PaymentContent = PaymentContent & "<input type=""hidden"" Name=""S_PostCode"" value=""" & s_postcode & """>" & vbNewLine                    '---- 消费者邮码
		PaymentContent = PaymentContent & "<input type=""hidden"" Name=""S_Telephone"" value=""" & s_tel & """>" & vbNewLine                        '---- 消费者电话
		PaymentContent = PaymentContent & "<input type=""hidden"" Name=""S_Email"" value=""" & s_eml & """>" & vbNewLine                            '---- 消费者邮件
		PaymentContent = PaymentContent & "<input type=""hidden"" Name=""R_Name"" value=""" & sConsignee & """>" & vbNewLine                        '---- 收货人姓名
		PaymentContent = PaymentContent & "<input type=""hidden"" Name=""R_Address"" value=""" & sAddress & """>" & vbNewLine                      '---- 收货人住址
		PaymentContent = PaymentContent & "<input type=""hidden"" Name=""R_PostCode"" value=""" & r_postcode & """>" & vbNewLine                    '---- 收货人邮码
		PaymentContent = PaymentContent & "<input type=""hidden"" Name=""R_Telephone"" value=""" & r_tel & """>" & vbNewLine                        '---- 收货人电话
		PaymentContent = PaymentContent & "<input type=""hidden"" Name=""R_Email"" value=""" & r_eml & """>" & vbNewLine                           '---- 收货人邮件
		PaymentContent = PaymentContent & "<input type=""hidden"" name=""MOComment"" value=""" & m_ocomment & """>" & vbNewLine                     '---- 备     注
		PaymentContent = PaymentContent & "<input type=""hidden"" Name=""MODate"" value=""" & modate & """>" & vbNewLine                            '---- 时间字段
		PaymentContent = PaymentContent & "<input type=""hidden"" Name=""State"" value=""" & sPstate & """>" & vbNewLine                            '---- 交易状态
		PaymentContent = PaymentContent & "<input type=""hidden"" Name=""digestinfo"" value=""" & digest & """>" & vbNewLine                        '---- 签名认证
		PaymentContent = PaymentContent & "<input Type=""submit"" Name=""submit"" value=""" & submit_value & """ class=""Button""> " & vbNewLine            '---- 确认支付
		PaymentContent = PaymentContent & "        </td>   </tr>" & vbNewLine
		PaymentContent = PaymentContent & "</form>" & vbNewLine
		PaymentContent = PaymentContent & "</table>" & vbNewLine
	End Sub
	'================================================
	'过程名：payment_chinabank
	'作  用：网银在线支付平台
	'================================================
	Private Sub Payment_chinabank()
		On Error Resume Next
		Dim v_mid, v_amount, v_oid, v_moneytype, style, v_url, remark1, remark2
		Dim OrderMessage, v_md5info
		Dim v_rcvname, v_rcvaddr, v_rcvtel, v_rcvpost, v_ordername, v_orderemail
		v_mid = sPaymentid                              '---- 商 家 号
		v_amount = sPaymoney                            '---- 金 额
		v_oid = Trim(sOrderid)                          '---- 定单号
		v_moneytype = sMoneytype                        '---- 币    种
		style = sPstate                                 '---- 指网关模式0(普通)，1(银行列表中带外卡)
		v_url = sReturnurl                              '---- 返回URL
		remark1 = sComment                              '---- 备 注1
		remark2 = sTelephone                               '---- 备 注2

		OrderMessage = v_amount & v_moneytype & v_oid & v_mid & v_url & sPaymentkey
		v_md5info = UCase(Trim(md5(OrderMessage,True)))                                                  '网银支付平台对MD5值只认大写字符串，所以小写的MD5值得转换为大写

		'**********以下几项与网上支付货款无关，建议不用**************
		v_rcvname = sConsignee                          '---- 收  货  人
		v_rcvaddr = sAddress                            '---- 收货人地址
		v_rcvtel = sTelephone                            '---- 收货人电话
		v_rcvpost = sTelephone                           '---- 消费者邮码
		v_ordername = sConsigner                        '---- 发  货  人
		v_orderemail = sEmail                           '---- 收货人邮件

		PaymentContent = "<table>         <tr>" & vbNewLine
		PaymentContent = PaymentContent & "<form method=""post"" action=""https://pay.chinabank.com.cn/select_bank"" name=""payform"" target=""_blank"">" & vbNewLine
		PaymentContent = PaymentContent & "        <td>" & vbNewLine
		PaymentContent = PaymentContent & "<input type=""hidden"" name=""v_md5info"" value=""" & v_md5info & """>" & vbNewLine
		PaymentContent = PaymentContent & "<input type=""hidden"" name=""v_mid"" value=""" & v_mid & """>" & vbNewLine
		PaymentContent = PaymentContent & "<input type=""hidden"" name=""v_oid"" value=""" & v_oid & """>" & vbNewLine
		PaymentContent = PaymentContent & "<input type=""hidden"" name=""v_amount"" value=""" & v_amount & """>" & vbNewLine
		PaymentContent = PaymentContent & "<input type=""hidden"" name=""v_moneytype""  value=""" & v_moneytype & """>" & vbNewLine
		PaymentContent = PaymentContent & "<input type=""hidden"" name=""v_url"" value=""" & v_url & """>" & vbNewLine
		PaymentContent = PaymentContent & "<input type=""hidden"" name=""style"" value=""" & style & """>" & vbNewLine
		PaymentContent = PaymentContent & "<input type=""hidden"" name=""remark1"" value=""" & remark1 & """>" & vbNewLine
		PaymentContent = PaymentContent & "<input type=""hidden"" name=""remark2"" value=""" & remark2 & """>" & vbNewLine
		'----- 以下几项与网上支付货款无关，建议不用 ----
		PaymentContent = PaymentContent & "<input type=""hidden"" name=""v_rcvname"" value=""" & v_rcvname & """>" & vbNewLine
		PaymentContent = PaymentContent & "<input type=""hidden"" name=""v_rcvaddr"" value=""" & v_rcvaddr & """>" & vbNewLine
		PaymentContent = PaymentContent & "<input type=""hidden"" name=""v_rcvtel"" value=""" & v_rcvtel & """>" & vbNewLine
		PaymentContent = PaymentContent & "<input type=""hidden"" name=""v_rcvpost"" value=""" & v_rcvpost & """>" & vbNewLine
		PaymentContent = PaymentContent & "<input type=""hidden"" name=""v_ordername""  value=""" & v_ordername & """>" & vbNewLine
		PaymentContent = PaymentContent & "<input type=""hidden"" name=""v_orderemail""  value=""" & v_orderemail & """>" & vbNewLine
		PaymentContent = PaymentContent & "<input type=""submit"" name=""v_action"" value=""" & submit_value & """ class=""Button"">" & vbNewLine
		PaymentContent = PaymentContent & "        </td>   </tr>" & vbNewLine
		PaymentContent = PaymentContent & "</form>" & vbNewLine
		PaymentContent = PaymentContent & "</table>" & vbNewLine
	End Sub
	'================================================
	'过程名：ReceivePage
	'作  用：在线支付返回页面
	'================================================
	Public Sub ReceivePage()
		On Error Resume Next
		If CInt(sPlatform) = 1 Then
			Call Receive_nps
		ElseIf CInt(sPlatform) = 2 Then
			Call Receive_chinabank
		Else
			ErrNumber = 6
			Exit Sub
		End If
	End Sub
	Private Sub Receive_nps()
		On Error Resume Next
		Dim OrderMessage, md5text, BankPayCode, md5info
		sPaymentid = Trim(Request("m_id"))                            '---- 商 家 号
		sOrderid = Trim(Request("m_orderid"))                         '---- 定 单 号
		sPaymoney = Trim(Request("m_oamount"))                        '---- 订单金额
		sComment = Trim(Request("m_ocomment"))                        '---- 备注
		sConsignee = Trim(Request("r_name"))                          '---- 收货人
		sAddress = Trim(Request("r_addr"))                            '---- 收货人地址
		sPostcode = Trim(Request("r_postcode"))                       '---- 收货人邮编
		sTelephone = Trim(Request("r_tel"))                           '---- 收货人电话
		sEmail = Trim(Request("r_eml"))                               '---- 收货人E-Mail
		sConsigner = Trim(Request("s_name"))                          '---- 发货人
		sPstate = Trim(Request("m_status"))                           '---- 交易状态
		sMoneytype = Trim(Request("m_ocurrency"))                     '---- 币种
		md5info = Trim(Request("newmd5info"))                         '---- 签名认证
		If Trim(Request("md5info")) = "" Then
			ErrNumber = 3
			Exit Sub
		End If
		OrderMessage = sPaymentid & sOrderid & sPaymoney & sPaymentkey & sPstate

		md5text = Trim(md5(OrderMessage,True))

		If UCase(md5text) <> UCase(md5info) Then
			ErrNumber = 4
			Exit Sub
		Else
			If ChkNumeric(sPstate) = 2 Then
				'---- 支付成功
				ErrNumber = 0
				Exit Sub
			Else
				ErrNumber = 4
				Exit Sub
			End If
		End If
	End Sub

	Private Sub Receive_chinabank()
		On Error Resume Next
		Dim v_md5str, md5text, OrderMessage, BankPayCode
		sPaymentid = sPaymentid                                       '---- 商 家 号
		sOrderid = Trim(Request("v_oid"))                             '---- 定 单 号
		sPaymoney = Trim(Request("v_amount"))                         '---- 订单金额
		sComment = Trim(Request("remark1"))                           '---- 备注
		sConsignee = Trim(Request("v_rcvname"))                       '---- 收货人
		sAddress = Trim(Request("v_rcvaddr"))                         '---- 收货人地址
		sPostcode = Trim(Request("v_rcvpost"))                        '---- 收货人邮编
		sTelephone = Trim(Request("remark2"))                         '---- 收货人电话
		sEmail = Trim(Request("v_orderemail"))                        '---- 收货人E-Mail
		sConsigner = Trim(Request("v_ordername"))                     '---- 发货人
		sMoneytype = Trim(Request("v_moneytype"))                     '---- 币种
		sPstate = Trim(Request("v_pstatus"))                          '---- 交易状态
		v_md5str = Trim(Request("v_md5str"))
		If Trim(Request("v_md5str")) = "" Then
			ErrNumber = 3
			Exit Sub
		End If
		OrderMessage = sOrderid & sPstate & sPaymoney & sMoneytype & sPaymentkey

		md5text = Trim(md5(OrderMessage,True))
		
		If UCase(md5text) <> UCase(v_md5str) Then
			ErrNumber = 4
			Exit Sub
		Else
			If ChkNumeric(sPstate) = 20 Then
				'---- 支付成功
				ErrNumber = 0
				Exit Sub
			Else
				'---- 支付失败
				ErrNumber = 4
				Exit Sub
			End If
		End If
	End Sub

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
		Dim MoneyPercent, valPercent
		
		MoneyPercent = mPercent / 100
		If MoneyPercent = 0 Then
			ReadPayMoney = money
		Else
			If CBool(Reduce) = True Then
				valPercent = Round(money / (1 + 1 * MoneyPercent), 2)
				ReadPayMoney = CCur(valPercent)
			Else
				valPercent = Round(money * MoneyPercent, 2)
				ReadPayMoney = CCur(money + valPercent)
			End If
		End If
	End Function

	Public Function ChkNumeric(ByVal CHECK_ID)
		If CHECK_ID <> "" And IsNumeric(CHECK_ID) Then _
			CHECK_ID = CLng(CHECK_ID) _
		Else _
			CHECK_ID = 0
		ChkNumeric = CHECK_ID
	End Function
End Class
%>
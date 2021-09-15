
<%
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
		sComment = "����֧��"
		sRemark = "����֧��"
		sConsigner = "Consigner"
		sConsignee = "Consignee"
		sAddress = "�˳�"
		sPostcode = "51800"
		sTelephone = "0359-8698845"
		sEmail = "liuyunfan@163.com"
		submit_value = "��������֧��ƽ̨"
		sEncrypt = "12345678"
		mPercent = 0
	End Sub

	Private Sub Class_Terminate()
		
	End Sub
	'---- �������
	Public Property Get Description()
		Select Case ErrNumber
			Case 1: Description = "�����Ŵ���"
			Case 2: Description = "���������"
			Case 3: Description = "��֤ǩ������Ϊ��ֵ!"
			Case 4: Description = "��֤��Ϣ�����˴ν���ʧ�ܣ�����"
			Case 5: Description = "��֤��Ϣ�����벻Ҫ�ظ��ύ���ݣ��˴ν���ʧ�ܣ�����"
			Case 6: Description = "�Բ��𣡱�վ��δ��ͨ����֧������,��ѡ����������֧��"
			Case 7: Description = "�����ϵͳ����"
			Case 8: Description = "��վ��δ��ͨ����֧������,���߱�վ����û��ע��,��ѡ����������֧��"
			Case Else
				Description = Empty
		End Select
	End Property
	'---- �����Ѱٷֱ�
	Public Property Let Percent(ByVal NewValue)
		mPercent = NewValue
	End Property
	'---- ֧��ƽ̨
	Public Property Let PayPlatform(ByVal NewValue)
		sPlatform = NewValue
	End Property
	'---- ֧��ID
	Public Property Let Paymentid(ByVal NewValue)
		sPaymentid = NewValue
		If Trim(sPaymentid) = "" Then
			sPaymentid = "1051011239"
		End If
	End Property
	Public Property Get Paymentid()
		Paymentid = sPaymentid
	End Property
	'---- ֧��KEY
	Public Property Let Paymentkey(ByVal NewValue)
		sPaymentkey = NewValue
		If Trim(sPaymentkey) = "" Then
			sPaymentkey = "enchiasp778899"
		End If
	End Property
	'---- ����URL
	Public Property Let Returnurl(ByVal NewValue)
		sReturnurl = NewValue
	End Property
	'---- ����״̬
	Public Property Let Pstate(ByVal NewValue)
		sPstate = NewValue
	End Property
	Public Property Get Pstate()
		Pstate = sPstate
	End Property
	'--- ������
	Public Property Let Orderid(ByVal NewValue)
		sOrderid = NewValue
	End Property
	Public Property Get Orderid()
		Orderid = sOrderid
	End Property
	'---- ֧�����
	Public Property Let Paymoney(ByVal NewValue)
		sPaymoney = ReadPayMoney(NewValue, False)
	End Property
	Public Property Get Paymoney()
		Paymoney = sPaymoney
	End Property
	'---- ���׽��
	Public Property Get Buymoney()
		Buymoney = ReadPayMoney(sPaymoney, True)
	End Property
	'---- ������
	Public Property Get ServiceCharge()
		ServiceCharge = sPaymoney - ReadPayMoney(sPaymoney, True)
	End Property
	'---- ֧������
	Public Property Let Moneytype(ByVal NewValue)
		sMoneytype = NewValue
	End Property
	Public Property Get Moneytype()
		Moneytype = sMoneytype
	End Property
	'---- ֧������
	Public Property Let Planguage(ByVal NewValue)
		sLanguage = NewValue
	End Property
	'---- ֧����ע
	Public Property Let Comment(ByVal NewValue)
		sComment = NewValue
	End Property
	Public Property Get Comment()
		Comment = sComment
	End Property
	'---- ֧����ע
	Public Property Let Remark(ByVal NewValue)
		sRemark = NewValue
	End Property
	Public Property Get Remark()
		Remark = sRemark
	End Property
	'---- �ջ�������
	Public Property Let Consignee(ByVal NewValue)
		sConsignee = NewValue
	End Property
	Public Property Get Consignee()
		Consignee = sConsignee
	End Property
	'---- �ջ��˵�ַ
	Public Property Let Address(ByVal NewValue)
		sAddress = NewValue
	End Property
	Public Property Get Address()
		Address = sAddress
	End Property
	'---- �ջ����ʱ�
	Public Property Let Postcode(ByVal NewValue)
		sPostcode = NewValue
	End Property
	Public Property Get Postcode()
		Postcode = sPostcode
	End Property
	'---- �ջ��˵绰
	Public Property Let Telephone(ByVal NewValue)
		sTelephone = NewValue
	End Property
	Public Property Get Telephone()
		Telephone = sTelephone
	End Property
	'---- �ջ���E_Mail
	Public Property Let Email(ByVal NewValue)
		sEmail = NewValue
	End Property
	Public Property Get Email()
		Email = sEmail
	End Property
	'---- ������
	Public Property Let Consigner(ByVal NewValue)
		sConsigner = NewValue
	End Property
	Public Property Get Consigner()
		Consigner = sConsigner
	End Property
	'---- �ύ��ť
	Public Property Let submitvalue(ByVal NewValue)
		submit_value = NewValue
	End Property
	'---- ��������
	Public Property Let Encrypt(ByVal NewValue)
		sEncrypt = NewValue
	End Property
	'================================================
	'��������GetWebSiteUrl
	'��  �ã�ȡ�ô��˿ڵ�URL
	'================================================
	Public Property Get GetWebSiteUrl()
		If Request.ServerVariables("SERVER_PORT") = "80" Then
			GetWebSiteUrl = "http://" & Request.ServerVariables("server_name")
		Else
			GetWebSiteUrl = "http://" & Request.ServerVariables("server_name") & ":" & Request.ServerVariables("SERVER_PORT")
		End If
	End Property
	'================================================
	'��������PaymentPlatform
	'��  �ã�����֧��ƽ̨
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
	'��������ShowPayment
	'��  �ã���ʾ����֧��ƽ̨
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
	'��������payment_nps
	'��  �ã�NPS����֧��ƽ̨
	'================================================
	Private Sub Payment_nps()
		On Error Resume Next
		Dim digest, OrderMessage
		Dim m_url, m_orderid, m_oamount, modate, m_ocomment
		Dim m_ocurrency, m_language, s_postcode, s_tel, s_eml, r_postcode, r_tel, r_eml
		m_orderid = Trim(sOrderid)                      '---- ������
		m_oamount = sPaymoney                           '---- �� ��
		m_url = sReturnurl                              '---- ����URL
		m_ocurrency = sMoneytype                        '---- ��    ��
		m_language = sLanguage                          '---- ����ѡ��
		s_postcode = sPostcode                          '---- ����������
		s_tel = sTelephone                              '---- ����������
		s_eml = sEmail                                  '---- �������ʼ�
		r_postcode = sPostcode                          '---- �ջ��˵绰
		r_tel = sTelephone                              '---- �ջ��˵绰
		r_eml = sEmail                                  '---- �ջ����ʼ�
		m_ocomment = sComment                           '---- �� ע
		modate = Date                                   '---- �� ��
		
		OrderMessage = sPaymentid & m_orderid & m_oamount & m_ocurrency & m_url & m_language & s_postcode & s_tel & s_eml & r_postcode & r_tel & r_eml & modate & sPaymentkey
		digest = UCase(Trim(md5(OrderMessage,True)))
		
		PaymentContent = "<table>         <tr>" & vbNewLine
		PaymentContent = PaymentContent & "<form method=""post"" action=""https://payment.nps.cn/VirReceiveMerchantAction.do"" name=""payform"" target=""_blank"">" & vbNewLine
		PaymentContent = PaymentContent & "        <td>" & vbNewLine
		PaymentContent = PaymentContent & "<input type=""hidden"" Name=""M_ID"" value=""" & sPaymentid & """>" & vbNewLine                         '---- �� �� ��
		PaymentContent = PaymentContent & "<input type=""hidden"" Name=""MOrderID"" value=""" & m_orderid & """>" & vbNewLine                       '---- �� �� ��
		PaymentContent = PaymentContent & "<input type=""hidden"" Name=""MOAmount"" value=""" & m_oamount & """>" & vbNewLine                       '---- �������
		PaymentContent = PaymentContent & "<input type=""hidden"" Name=""MOCurrency"" value=""" & m_ocurrency & """>" & vbNewLine                   '---- ��    ��
		PaymentContent = PaymentContent & "<input type=""hidden"" name=""M_URL"" value=""" & m_url & """>" & vbNewLine                              '---- ���ص�ַ
		PaymentContent = PaymentContent & "<input type=""hidden"" Name=""M_Language"" value=""" & m_language & """>" & vbNewLine                    '---- ����ѡ��
		PaymentContent = PaymentContent & "<input type=""hidden"" Name=""S_Name"" value=""" & sConsignee & """>" & vbNewLine                        '---- ����������
		PaymentContent = PaymentContent & "<input type=""hidden"" Name=""S_Address"" value=""" & sAddress & """>" & vbNewLine                       '---- ������סַ
		PaymentContent = PaymentContent & "<input type=""hidden"" Name=""S_PostCode"" value=""" & s_postcode & """>" & vbNewLine                    '---- ����������
		PaymentContent = PaymentContent & "<input type=""hidden"" Name=""S_Telephone"" value=""" & s_tel & """>" & vbNewLine                        '---- �����ߵ绰
		PaymentContent = PaymentContent & "<input type=""hidden"" Name=""S_Email"" value=""" & s_eml & """>" & vbNewLine                            '---- �������ʼ�
		PaymentContent = PaymentContent & "<input type=""hidden"" Name=""R_Name"" value=""" & sConsignee & """>" & vbNewLine                        '---- �ջ�������
		PaymentContent = PaymentContent & "<input type=""hidden"" Name=""R_Address"" value=""" & sAddress & """>" & vbNewLine                      '---- �ջ���סַ
		PaymentContent = PaymentContent & "<input type=""hidden"" Name=""R_PostCode"" value=""" & r_postcode & """>" & vbNewLine                    '---- �ջ�������
		PaymentContent = PaymentContent & "<input type=""hidden"" Name=""R_Telephone"" value=""" & r_tel & """>" & vbNewLine                        '---- �ջ��˵绰
		PaymentContent = PaymentContent & "<input type=""hidden"" Name=""R_Email"" value=""" & r_eml & """>" & vbNewLine                           '---- �ջ����ʼ�
		PaymentContent = PaymentContent & "<input type=""hidden"" name=""MOComment"" value=""" & m_ocomment & """>" & vbNewLine                     '---- ��     ע
		PaymentContent = PaymentContent & "<input type=""hidden"" Name=""MODate"" value=""" & modate & """>" & vbNewLine                            '---- ʱ���ֶ�
		PaymentContent = PaymentContent & "<input type=""hidden"" Name=""State"" value=""" & sPstate & """>" & vbNewLine                            '---- ����״̬
		PaymentContent = PaymentContent & "<input type=""hidden"" Name=""digestinfo"" value=""" & digest & """>" & vbNewLine                        '---- ǩ����֤
		PaymentContent = PaymentContent & "<input Type=""submit"" Name=""submit"" value=""" & submit_value & """ class=""Button""> " & vbNewLine            '---- ȷ��֧��
		PaymentContent = PaymentContent & "        </td>   </tr>" & vbNewLine
		PaymentContent = PaymentContent & "</form>" & vbNewLine
		PaymentContent = PaymentContent & "</table>" & vbNewLine
	End Sub
	'================================================
	'��������payment_chinabank
	'��  �ã���������֧��ƽ̨
	'================================================
	Private Sub Payment_chinabank()
		On Error Resume Next
		Dim v_mid, v_amount, v_oid, v_moneytype, style, v_url, remark1, remark2
		Dim OrderMessage, v_md5info
		Dim v_rcvname, v_rcvaddr, v_rcvtel, v_rcvpost, v_ordername, v_orderemail
		v_mid = sPaymentid                              '---- �� �� ��
		v_amount = sPaymoney                            '---- �� ��
		v_oid = Trim(sOrderid)                          '---- ������
		v_moneytype = sMoneytype                        '---- ��    ��
		style = sPstate                                 '---- ָ����ģʽ0(��ͨ)��1(�����б��д��⿨)
		v_url = sReturnurl                              '---- ����URL
		remark1 = sComment                              '---- �� ע1
		remark2 = sTelephone                               '---- �� ע2

		OrderMessage = v_amount & v_moneytype & v_oid & v_mid & v_url & sPaymentkey
		v_md5info = UCase(Trim(md5(OrderMessage,True)))                                                  '����֧��ƽ̨��MD5ֵֻ�ϴ�д�ַ���������Сд��MD5ֵ��ת��Ϊ��д

		'**********���¼���������֧�������޹أ����鲻��**************
		v_rcvname = sConsignee                          '---- ��  ��  ��
		v_rcvaddr = sAddress                            '---- �ջ��˵�ַ
		v_rcvtel = sTelephone                            '---- �ջ��˵绰
		v_rcvpost = sTelephone                           '---- ����������
		v_ordername = sConsigner                        '---- ��  ��  ��
		v_orderemail = sEmail                           '---- �ջ����ʼ�

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
		'----- ���¼���������֧�������޹أ����鲻�� ----
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
	'��������ReceivePage
	'��  �ã�����֧������ҳ��
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
		sPaymentid = Trim(Request("m_id"))                            '---- �� �� ��
		sOrderid = Trim(Request("m_orderid"))                         '---- �� �� ��
		sPaymoney = Trim(Request("m_oamount"))                        '---- �������
		sComment = Trim(Request("m_ocomment"))                        '---- ��ע
		sConsignee = Trim(Request("r_name"))                          '---- �ջ���
		sAddress = Trim(Request("r_addr"))                            '---- �ջ��˵�ַ
		sPostcode = Trim(Request("r_postcode"))                       '---- �ջ����ʱ�
		sTelephone = Trim(Request("r_tel"))                           '---- �ջ��˵绰
		sEmail = Trim(Request("r_eml"))                               '---- �ջ���E-Mail
		sConsigner = Trim(Request("s_name"))                          '---- ������
		sPstate = Trim(Request("m_status"))                           '---- ����״̬
		sMoneytype = Trim(Request("m_ocurrency"))                     '---- ����
		md5info = Trim(Request("newmd5info"))                         '---- ǩ����֤
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
				'---- ֧���ɹ�
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
		sPaymentid = sPaymentid                                       '---- �� �� ��
		sOrderid = Trim(Request("v_oid"))                             '---- �� �� ��
		sPaymoney = Trim(Request("v_amount"))                         '---- �������
		sComment = Trim(Request("remark1"))                           '---- ��ע
		sConsignee = Trim(Request("v_rcvname"))                       '---- �ջ���
		sAddress = Trim(Request("v_rcvaddr"))                         '---- �ջ��˵�ַ
		sPostcode = Trim(Request("v_rcvpost"))                        '---- �ջ����ʱ�
		sTelephone = Trim(Request("remark2"))                         '---- �ջ��˵绰
		sEmail = Trim(Request("v_orderemail"))                        '---- �ջ���E-Mail
		sConsigner = Trim(Request("v_ordername"))                     '---- ������
		sMoneytype = Trim(Request("v_moneytype"))                     '---- ����
		sPstate = Trim(Request("v_pstatus"))                          '---- ����״̬
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
				'---- ֧���ɹ�
				ErrNumber = 0
				Exit Sub
			Else
				'---- ֧��ʧ��
				ErrNumber = 4
				Exit Sub
			End If
		End If
	End Sub

	'=============================================================
	'��������ReadPayMoney
	'��  �ã���ȡҪ֧���Ľ�Ǯ
	'��  ����money   ----ʵ�ʽ�Ǯ
	'����ֵ�����������Ѻ�Ľ�Ǯ
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
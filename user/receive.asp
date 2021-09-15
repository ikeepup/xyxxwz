<!--#include file="config.asp"-->
<!--#include file="check.asp"-->
<!--#include file="../inc/cls_payment.asp"-->
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
Dim m_orderid,addmoney,m_oamount,m_ocomment
Dim wp,strChinaeBank
strChinaeBank = Split(enchiasp.ChinaeBank, "|||")
Set wp = New WebPayment_Cls
wp.PayPlatform = CInt(enchiasp.StopBankPay)
wp.Paymentid = Trim(strChinaeBank(0))
wp.Paymentkey = Trim(strChinaeBank(1))
wp.Percent = enchiasp.CheckNumeric(strChinaeBank(2))
wp.Comment = "会员充值"
wp.ReceivePage
m_orderid = enchiasp.CheckInfuse(wp.Orderid,30)
addmoney = wp.Buymoney
m_oamount = wp.Paymoney
m_ocomment = wp.Comment
Select Case CInt(wp.ErrNumber)
Case 0
	SaveUserInfo m_orderid,addmoney,m_oamount,m_ocomment
Case 3
	ErrMsg = wp.Description
	Founderr = True
Case 4
	ErrMsg = wp.Description
	Founderr = True
Case 5
	ErrMsg = wp.Description
	Founderr = True
Case 6
	ErrMsg = wp.Description
	Founderr = True
Case 8
	ErrMsg = wp.Description
	Founderr = True
End Select
Set wp = Nothing
If Founderr = True Then
	Call Returnerr(ErrMsg)
End If
CloseConn
Function SaveUserInfo(OrderForm,addmoney,realmoney,readme)
	Dim Rs,SQL
	Set Rs = enchiasp.Execute("SELECT id FROM ECCMS_AddMoney WHERE OrderForm='"& enchiasp.CheckStr(OrderForm) &"'")
	If Not (Rs.BOF And Rs.EOF) Then
		ErrMsg = ErrMsg + "<li>认证信息出错，请不要重复提交数据,此次交易失败！！！</li>"
		Founderr = True
		Set Rs = Nothing
		Exit Function
	End If
	Set Rs = Nothing
	If Founderr = True Then Exit Function
	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL = "SELECT * FROM ECCMS_AddMoney WHERE (id is null)"
	Rs.Open SQL,Conn,1,3
	Rs.AddNew
		Rs("userid").Value = enchiasp.memberid
		Rs("username").Value = enchiasp.membername
		Rs("title").Value = enchiasp.ChkFormStr(readme)
		Rs("OrderForm").Value = Trim(OrderForm)
		Rs("addmoney").Value = CCur(realmoney)
		Rs("addtime").Value = Now()
		Rs("readme").Value = enchiasp.ChkFormStr(readme)
		Rs("paytype").Value = "在线支付"
		Rs("finished").Value = 1
		Rs("deletion").Value = 0
	Rs.Update
	Rs.Close:Set Rs = Nothing
	
	enchiasp.Execute ("UPDATE ECCMS_User SET usermoney=usermoney+"& CCur(addmoney) &" WHERE username='"& enchiasp.CheckRequest(enchiasp.membername,50) &"' And userid="& CLng(enchiasp.memberid))
	
	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL = "SELECT * FROM ECCMS_Account WHERE (AccountID is null)"
	Rs.Open SQL,Conn,1,3
	Rs.addnew
		Rs("payer").Value = enchiasp.membername
		Rs("payee").Value = enchiasp.SiteName
		Rs("product").Value = enchiasp.ChkFormStr(readme)
		Rs("Amount").Value = 1
		Rs("unit").Value = "次"
		Rs("price").Value = CCur(addmoney)
		Rs("TotalPrices").Value = CCur(realmoney)
		Rs("DateAndTime").Value = Now()
		Rs("Accountype").Value = 0
		Rs("Explain").Value = enchiasp.ChkFormStr(readme)
		Rs("Reclaim").Value = 0
	Rs.update
	Rs.Close:Set Rs = Nothing
	Call Returnsuc("<li>恭喜您！认证成功，会员充值完成。</li><li>实收金额："& FormatCurrency(realmoney,2,-1) &" 元</li><li>充值金额："& FormatCurrency(addmoney,2,-1) &" 元</li><li>此次所用手续费："& FormatCurrency(realmoney-addmoney,2,-1) &" 元</li>")
End Function
%>

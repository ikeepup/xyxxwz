function CheckForm(form1)
{
	if (!validateSubmit()) return (false);
	form1.content.value=getHTML(); 
	MessageLength=IframeID.document.body.innerHTML.length;
	if(MessageLength<2){alert("商品说明不能小于2个字符！");return false;}
	if (form1.TradeName.value == "")
	{
		alert("商品名称不能为空！");
		form1.TradeName.focus();
		return (false);
	}
	if (form1.ClassID.value==""){
		alert("该一级分类已经有下属分类，请选择其下属分类！");
		form1.ClassID.focus();
		return (false);
	}
	if (form1.ClassID.value=="0"){
		alert("该分类是外部连接，不能添加内容！");
		form1.ClassID.focus();
		return (false);
	}
	if (form1.Unit.value==""){
		alert("请选择商品单位！");
		form1.Unit.focus();
		return false;
	}
	if (form1.PastPrice.value==""){
		alert("请输入商品原价！");
		form1.PastPrice.focus();
		return false;
	}
	if (form1.NowPrice.value==""){
		alert("请输入商品价格！");
		form1.NowPrice.focus();
		return false;
	}
	//document.myform.submit1.disabled = true;
}
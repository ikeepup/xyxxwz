function CheckForm(myform){
	myform.content.value=getHTML(); 
	MessageLength=IframeID.document.body.innerHTML.length;
	if(MessageLength<LeastString){alert("留言内容不能小于"+LeastString+"个字符！");return false;}
	if(MessageLength>MaxString){alert("留言内容不能大于"+MaxString+"个字符！");return false;}
	if (myform.username.value==""){
		alert("用户名称不能为空！");
		document.myform.username.focus();
		return false;
	}
	if (myform.topic.value==""){
		alert("留言主题不能为空！");
		document.myform.topic.focus();
		return false;
	}
	if ((myform.GuestEmail.value.indexOf("@") == -1) || (myform.GuestEmail.value.indexOf(".") == -1)){
		alert("请查看您的E-mail地址是否正确，请重录入!");
		document.myform.GuestEmail.focus();
       		return false;
	}
	
	//document.myform.submit1.disabled = true;
	//document.myform.submit();
}

function formatbt()
{
  var arr = showModalDialog("../editor/btformat.htm?",null, "dialogWidth:250pt;dialogHeight:166pt;toolbar=no;location=no;directories=no;status=no;menubar=NO;scrollbars=no;resizable=no;help=0; status:0");
  if (arr != null){
     document.myform.Topicformat.value=arr;
     myt.innerHTML="<span style='background-color: #FFFFff;font-size:12px' "+arr+">设置标题样式 ABCdef</span>";
  }
}
function Cancelform()
{
  document.myform.Topicformat.value='';
  myt.innerHTML="<span style='background-color: #FFFFff;font-size:12px'>设置标题样式 ABCdef</span>";
}
function CtrlEnter()
{
	if(event.ctrlKey && window.event.keyCode==13)
	{
		this.document.myform.submit();
	}	
}
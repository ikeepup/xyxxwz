function CheckForm(myform){
	myform.content.value=getHTML(); 
	MessageLength=IframeID.document.body.innerHTML.length;
	if(MessageLength<LeastString){alert("�������ݲ���С��"+LeastString+"���ַ���");return false;}
	if(MessageLength>MaxString){alert("�������ݲ��ܴ���"+MaxString+"���ַ���");return false;}
	if (myform.username.value==""){
		alert("�û����Ʋ���Ϊ�գ�");
		document.myform.username.focus();
		return false;
	}
	if (myform.topic.value==""){
		alert("�������ⲻ��Ϊ�գ�");
		document.myform.topic.focus();
		return false;
	}
	if ((myform.GuestEmail.value.indexOf("@") == -1) || (myform.GuestEmail.value.indexOf(".") == -1)){
		alert("��鿴����E-mail��ַ�Ƿ���ȷ������¼��!");
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
     myt.innerHTML="<span style='background-color: #FFFFff;font-size:12px' "+arr+">���ñ�����ʽ ABCdef</span>";
  }
}
function Cancelform()
{
  document.myform.Topicformat.value='';
  myt.innerHTML="<span style='background-color: #FFFFff;font-size:12px'>���ñ�����ʽ ABCdef</span>";
}
function CtrlEnter()
{
	if(event.ctrlKey && window.event.keyCode==13)
	{
		this.document.myform.submit();
	}	
}
function CheckAll(form) {  
	for (var i=0;i<form.elements.length;i++)  {  
		var e = form.elements[i];  
		if (e.name != 'chkall')  
		e.checked = true // form.chkall.checked;  
	}  
} 
 
function ContraSel(form) {
	for (var i=0;i<form.elements.length;i++)  {
		var e = form.elements[i];
		if (e.name != 'chkall')
		e.checked=!e.checked;
	}
}
function CheckAll2(form)  
{  
  	for (var i=0;i<form.elements.length;i++)  {  
    		var e = form.elements[i];  
    		if (e.name != 'chkall')  
       		e.checked = form.chkall.checked;  
    	}  
}
function CtrlEnter()
{
	if(event.ctrlKey && window.event.keyCode==13)
	{
		this.document.myform.submit();
	}	
}
function openDialog(url, width, height){
	var Win = window.showModalDialog(url,'openDialog','dialogWidth:' + width + 'px;dialogHeight:' + height + 'px;help:no;scroll:no;status:no');
}
function showClick(msg){
	if(confirm(msg)){
		event.returnValue=true;
	}else{
		event.returnValue=false;
	}
}
function showsub(ss)
{
ss=document.getElementById(ss)
 if (ss.style.display=="none") 
  {ss.style.display="";

}
 else
  {ss.style.display="none";
   }
}

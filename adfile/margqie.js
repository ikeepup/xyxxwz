marqueesHeight=520;
stopscroll=false;

with(marquees){
  style.width=228;
  style.height=marqueesHeight;
  style.overflowX="hidden";
  style.overflowY="hidden";
  noWrap=true;
  onmouseover=new Function("stopscroll=true");
  onmouseout=new Function("stopscroll=false");
}
document.write('<div id="templayer" style="position:absolute;z-index:1;visibility:hidden;"></div>');

preTop=0; currentTop=0; 

function init(){
  templayer.innerHTML="";
  while(templayer.offsetHeight<marqueesHeight){
    templayer.innerHTML+=marquees.innerHTML;
  }
  marquees.innerHTML=templayer.innerHTML+templayer.innerHTML;
  setInterval("scrollUp()",20);//Ô½´óÔ½Âý
}
document.body.onload=init;

function scrollUp(){
  if(stopscroll==true) return;
  preTop=marquees.scrollTop;
  marquees.scrollTop+=1;
  if(preTop==marquees.scrollTop){
    marquees.scrollTop=templayer.offsetHeight-marqueesHeight;
    marquees.scrollTop+=1;
  }
}

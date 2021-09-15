<%
dim Action 
Action = LCase(Request("id"))

select case Action

case "0"

Response.Redirect("../index.asp")

case "1"

Response.Redirect("../about/list.asp?classid=161")
case "2"

Response.Redirect("../about/list.asp?classid=162")
case "3"

Response.Redirect("../about/list.asp?classid=163")
case "4"

Response.Redirect("../about/list.asp?classid=164")
case "5"

Response.Redirect("../about/list.asp?classid=165")
case "6"

Response.Redirect("../about/list.asp?classid=166")


case "7"

Response.Redirect("../jiameng/list.asp?classid=179")
case "8"

Response.Redirect("../jiameng/list.asp?classid=180")
case "9"

Response.Redirect("../jiameng/list.asp?classid=181")
case "10"

Response.Redirect("../jiameng/list.asp?classid=182")

case "11"

Response.Redirect("../jiameng/list.asp?classid=183")
case "12"

Response.Redirect("../jiameng/list.asp?classid=184")
case "13"

Response.Redirect("../jiameng/list.asp?classid=185")
case "14"

Response.Redirect("../jiameng/list.asp?classid=186")
case "15"

Response.Redirect("../jiameng/list.asp?classid=187")
case "16"

Response.Redirect("../shop/")

case "17"

Response.Redirect("../shop/showbest.asp")
case "18"

Response.Redirect("../shop/shownew.asp")
case "19"

Response.Redirect("../article/list.asp?classid=169")
case "20"

Response.Redirect("../zxsc-gwzn/list.asp?classid=347")
case "21"

Response.Redirect("../article/list.asp?classid=167")
case "22"

Response.Redirect("../article/list.asp?classid=168")
case "23"

Response.Redirect("../article/list.asp?classid=169")
case "24"

Response.Redirect("../article/list.asp?classid=170")

case "25"

Response.Redirect("../article/list.asp?classid=171")





case "26"

Response.Redirect("../peixun/list.asp?classid=193")
case "27"

Response.Redirect("../peixun/list.asp?classid=194")
case "28"

Response.Redirect("../peixun/list.asp?classid=195")
case "29"

Response.Redirect("../peixun/list.asp?classid=196")
case "30"

Response.Redirect("../peixun/list.asp?classid=197")





case "31"

Response.Redirect("../shequ/list.asp?classid=188")
case "32"

Response.Redirect("../shequ/list.asp?classid=189")
case "33"

Response.Redirect("../shequ/list.asp?classid=190")
case "34"

Response.Redirect("../shequ/list.asp?classid=191")


case "35"

Response.Redirect("../about/list.asp?classid=192")


case "36"

Response.Redirect("../zhiying/list.asp?classid=369")

case "37"

Response.Redirect("../others/list.asp?classid=371")

case "38"

Response.Redirect("../others/list.asp?classid=373")

case "39"

Response.Redirect("../zhiying/list.asp?classid=374")

case "40"

Response.Redirect("../zhiying/list.asp?classid=375")

case else

Response.Redirect("../index.asp")
end select
%>

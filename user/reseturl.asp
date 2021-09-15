<%
dim Action 
Action = LCase(Request("id"))

select case Action

case "0"

Response.Redirect("../index_gb.asp")
case "01"

Response.Redirect("../about")
case "1"

Response.Redirect("../about")
case "2"

Response.Redirect("../about/list.asp?classid=19")
case "3"

Response.Redirect("../about/list.asp?classid=20")
case "4"

Response.Redirect("../about/list.asp?classid=21")
case "5"

Response.Redirect("../about/list.asp?classid=22")
case "6"

Response.Redirect("../about/list.asp?classid=23")
case "07"

Response.Redirect("../article")
case "7"

Response.Redirect("../article/list.asp?classid=1")
case "8"

Response.Redirect("../article/list.asp?classid=10")
case "9"

Response.Redirect("../article/list.asp?classid=11")
case "10"

Response.Redirect("../article/list.asp?classid=12")
case "011"

Response.Redirect("../ywjs")
case "11"

Response.Redirect("../ywjs/list.asp?classid=32")
case "12"

Response.Redirect("../ywjs/list.asp?classid=33")
case "13"

Response.Redirect("../ywjs/list.asp?classid=34")
case "14"

Response.Redirect("../ywjs/list.asp?classid=35")
case "15"

Response.Redirect("../ywjs/list.asp?classid=36")
case "16"

Response.Redirect("../ywjs/list.asp?classid=37")
case "017"

Response.Redirect("../fwzx")
case "17"

Response.Redirect("../fwzx/list.asp?classid=38")
case "18"

Response.Redirect("../fwzx/list.asp?classid=39")
case "19"

Response.Redirect("../fwzx/list.asp?classid=40")
case "20"

Response.Redirect("../fwzx/list.asp?classid=41")
case "21"

Response.Redirect("../fwzx/list.asp?classid=42")
case "22"

Response.Redirect("../fwzx/list.asp?classid=43")
case "23"

Response.Redirect("../fwzx/list.asp?classid=44")
case "24"

Response.Redirect("../fwzx/list.asp?classid=45")
case "025"

Response.Redirect("../cpzs")
case "25"

Response.Redirect("../cpzs/list.asp?classid=46")
case "26"

Response.Redirect("../cpzs/list.asp?classid=47")
case "27"

Response.Redirect("../cpzs/list.asp?classid=48")
case "28"

Response.Redirect("../cpzs/list.asp?classid=49")
case "29"

Response.Redirect("../cpzs/list.asp?classid=50")
case "30"

Response.Redirect("../cpzs/list.asp?classid=51")
case "031"

Response.Redirect("../job")
case "31"

Response.Redirect("../job/list.asp?classid=52")
case "32"

Response.Redirect("../job/list.asp?classid=53")
case "33"

Response.Redirect("../job/list.asp?classid=54")
case "34"

Response.Redirect("../job/list.asp?classid=55")
case "35"

Response.Redirect("../job/list.asp?classid=56")
case "36"

Response.Redirect("../shop")
case "37"

Response.Redirect("../GuestBook")
case "38"

Response.Redirect("../index_gb.asp")
case "39"

Response.Redirect("../tdqk/list.asp?classid=109")
case "40"

Response.Redirect("../link")
case "41"

Response.Redirect("../guowu/list.asp?classid=105")
case "42"

Response.Redirect("../guowu/list.asp?classid=106")
case "43"

Response.Redirect("../user/login.asp")
case "45"

Response.Redirect("../guowu/list.asp?classid=107")
case "46"

Response.Redirect("../gwzx")
case "47"

Response.Redirect("../guowu/list.asp?classid=108")
case else

Response.Redirect("../index_gb.asp")
end select
%>

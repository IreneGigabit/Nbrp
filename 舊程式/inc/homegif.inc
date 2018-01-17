<%
function homegif(a,b)
   Set conn = Server.CreateObject("ADODB.Connection")
   conn.Open Session("ODBCDSN")
   strsql="select homegif from logingrp where syscode='"_
         & trim(a) &"' and logingrp = '" & trim(b) &"'"
   set rs=conn.Execute(strsql)
   if rs.eof then
      homegif="trade.htm"
   else
      homegif=trim(rs("homegif"))
   end if   
   rs.close
   set rs=nothing
   conn.close
   set conn=nothing
end function   
%>   
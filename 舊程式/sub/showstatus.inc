<script language="javascript">
var TC=0;
var msg0="                        ";
var msg1="                        ";
var msg2="                        ";
var msg3="                        ";
var msg4="資料抓取中，請稍後......";
var msg=msg0+msg1+msg2+msg3+msg4
var timer=0;
function mess()
{
	if (TC<1000) 
	{	TC++;
		window.status = msg;
		msg = msg.substring(2,msg.length)+msg.substring(0,2);
		var timeID = setTimeout("",1);
		clearTimeout(timeID);
		timeID=timer+setTimeout("mess()",100); 
	}
	else
	{	window.status = "";}
}
mess();
</script>
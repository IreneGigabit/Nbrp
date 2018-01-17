<script Language=javaScript>
function chdisplayall(pobject,pno)
{
	if (pobject.OpenAll.value=="场甶秨") {
		for (i=1;i<=pno;i++) {
			if (pobject.all("id"+i).style.display=="none") {
				pobject.all("I"+i).src="../sub/2.gif";
				pobject.all("id"+i).style.display="";
			}
		}
		pobject.OpenAll.value="场闽超";
	}
	else
	{
		if (pobject.OpenAll.value=="场闽超")
		{
			for (i=1;i<=pno;i++) {
				if (pobject.all("id"+i).style.display=="") {
					pobject.all("I"+i).src="../sub/1.gif";
					pobject.all("id"+i).style.display="none";
				}
			}
			pobject.OpenAll.value="场甶秨";
		}
	}
}
function chdisplay(pobject,k)
{
	if (pobject.all("id"+k).style.display=="")
	{
		pobject.all("I"+k).src="../sub/1.gif";
		pobject.all("id"+k).style.display="none";
	}
	else
	{
		pobject.all("I"+k).src="../sub/2.gif";
		pobject.all("id"+k).style.display="";
	}
}
</script>
<script Language=javaScript>
function chdisplayall(pobject,pno)
{
	if (pobject.OpenAll.value=="�����i�}") {
		for (i=1;i<=pno;i++) {
			if (pobject.all("id"+i).style.display=="none") {
				pobject.all("I"+i).src="../sub/2.gif";
				pobject.all("id"+i).style.display="";
			}
		}
		pobject.OpenAll.value="��������";
	}
	else
	{
		if (pobject.OpenAll.value=="��������")
		{
			for (i=1;i<=pno;i++) {
				if (pobject.all("id"+i).style.display=="") {
					pobject.all("I"+i).src="../sub/1.gif";
					pobject.all("id"+i).style.display="none";
				}
			}
			pobject.OpenAll.value="�����i�}";
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
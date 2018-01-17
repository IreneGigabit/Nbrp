	<TR>
		<td class="greentext" align="right" id=salename>Àç&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;¬¢ :</td>
		<td class="greendata" align="left" colspan=3>
		<%If (HTProgRight AND 64) <> 0 then%>
			<select name="scode" size="1">
				<option value="all">¥þ³¡</option>
				<%SQL = "SELECT scode, sc_name FROM vcust_scode "
				SQL = SQL & " WHERE dept = 'P' AND  Syscode = '" & session("Syscode") & "' and branch ='" & session("se_branch") & "' and worktype = 'sales'  order by scode"
				call ShowSelect(cnn,SQL,true)%>
			</SELECT>		
		<%Else %>
			<input type="text" name="scode" readonly class="gSEdit" size=5 value="<%=session("se_scode")%>">-<%=session("sc_name")%>
		<%End if %>	
		</td>
	</TR>

<%
Response.Write "<div class=""module-box box-blue icon-selection""><div><div><div>"
if Request.QueryString("sub") = "add" then
	add_vote()
else

function ipadres(degis)
degis = Replace (degis, "''", "", 1, -1, 1)
ipadres=degis
end function

	SQL = "SELECT * FROM gop_anketsoru, gop_anketcevap WHERE gop_anketsoru.active = True AND gop_anketsoru.id = gop_anketcevap .poll_id order by gop_anketcevap.no_votes desc"
	Set rs = baglanti.execute(SQL)
	
	if not rs.EOF then
		poll_id = rs("id")
		all_voters = rs("votes")
		expir_date = rs("expiration_end")
		start_date = rs("expiration_start")
		SQL_IP = "SELECT * FROM anket_ip WHERE poll_id_ip=" & poll_id & " AND ip='" & ipadres(Request.ServerVariables("REMOTE_ADDR")) & "'"	 
		Set rs_IP = baglanti.Execute(SQL_IP)
		
		
	end if
	
cookie_id = Request.Cookies("currpoll")

Response.Write "<table border=""0"" cellspacing=""0"" cellpadding=""2""><tr><td><h3>"& polls &"</h3></td>  </tr>"
  
  if rs.eof or expir_date < date or start_date > date then 

Response.Write "<tr><td><table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""2""><tr><td align=""center"">"& no_active_poll &"</td></tr></table></td>	</tr>"
  
 else
  
		if not rs_IP.eof or cookie_id = Cstr(poll_id) then
	
Response.Write "<tr><td><table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""2""><tr><td align=""left"">"&rs.Fields("title")&"</td></tr><tr><td align=""left"">"

					do			
						b = Clng(rs.fields("no_votes"))
						if b = "0" then
Response.Write "<b>"&rs.fields("answer")&"</b>&nbsp;&nbsp;0%<br />"& no_vote &"<br />"					
					 	else
					 c = Clng(100 / all_voters * b)			
Response.Write "<b>"&rs.fields("answer")&"</b>&nbsp;&nbsp;"& c & "%" &"<br /><img src=""images/bar.gif"" height=""6"" width="""& 1*c &""" alt="""&rs.Fields("no_votes")&"""><br />"
						end if
				
					rs.movenext
					loop while not rs.eof
Response.Write "<br />"& total_votes &": <font color=""red""><b>"&all_voters&"</b></font></td></tr></table></td></tr>"
else

Response.Write "<tr><td><table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""2""><form name=""addVote"" method=""POST"" action=""default.asp?sub=add&id="&poll_id&"""><tr>  <td colspan=""2"" align=""left"">"&rs.Fields("title")&"</td>		      </tr>"
		      
				do 
Response.Write "<tr>   <td width=""20""><input type=""radio"" name=""voteFor"" value="""&rs.Fields("answer_id")&""" /></td><td align=""left"">"&rs.Fields("answer")&"</td></tr>"
		      		      	
				rs.MoveNext
				loop until rs.EOF
				
Response.Write "<tr><td colspan=""2"" align=""center""><input type=""submit"" value="""& vote &""" /></td></tr></form></table></td></tr>"
		
end if
	
end if
  
  if expir_date >= date and not start_date > date then
  calc = expir_date - 5
  if calc <= date then


Response.Write "<tr><td align=""center""> "
	
		select case cstr(date - calc)
		case "5":
			Response.Write last_days
		case "4":
			Response.Write last_1_days
		case "3":
			Response.Write last_2_days
		case "2":
			Response.Write last_3_days
		case "1":
			Response.Write last_4_days
		case "0":
			Response.Write last_5_days
		end select
		
Response.Write "</td></tr>"
  
  end if
  end if

		
		rs.close
		set rs = nothing
			
		SQL_inact = "SELECT * FROM gop_anketsoru WHERE active = False ORDER BY id DESC"
		set rs = baglanti.Execute(SQL_inact)
						
		a = 1

		if rs.eof then
		
		
Response.Write ""
	
		else
			if  rs("expiration_start") > date then
			
Response.Write ""
else

Response.Write ""

do 


Response.Write ""

					rs.movenext
					a = a + 1
					loop until rs.eof

Response.Write "</div> </td></tr>"		

			end if
		end if
	

Response.Write "</table>"



rs.close
set rs = nothing

end if

sub add_vote()

on error resume next
	
	poll_id = Request.QueryString("id")
	cookie_id = Request.Cookies("currpoll")

	SQL_IP = "SELECT * FROM anket_ip WHERE poll_id_ip=" & poll_id & " AND ip='" & ipadres(Request.ServerVariables("REMOTE_ADDR")) & "'"	 
	Set rs_IP = baglanti.Execute(SQL_IP)
	
	If cookie_id = Cstr(poll_id) or Request.form("voteFor") = "" or not rs_IP.Eof then
		
		Response.Redirect(Request.ServerVariables("URL"))
	
	else
		
		SQL_upd = "UPDATE gop_anketcevap SET no_votes=no_votes + 1 WHERE answer_id=" & int(Request.form("voteFor"))
		SQL_no_votes = "UPDATE gop_anketsoru SET votes=votes + 1 WHERE id=" & int(poll_id)
		SQL_expir = "SELECT * FROM gop_anketsoru WHERE id=" & int(poll_id)
		SQL_ip_block = "INSERT INTO anket_ip (poll_id_ip, ip) VALUES (" & int(poll_id) & ",'" & "''"& Request.ServerVariables("REMOTE_ADDR") & "''"& "')"
		set rs_upd = baglanti.execute(SQL_upd)
		set rs_no_votes = baglanti.execute(SQL_no_votes)
		set rs_expir = baglanti.execute(SQL_expir)
		set rs_ip_block = baglanti.execute(SQL_ip_block)
		
			Response.cookies("currpoll") = poll_id
			Response.Cookies("currpoll").Expires = rs_expir("expiration")
		
			rs_upd.Close
			rs_no_votes.close
			rs_expir.close
			rs_ip_blokck.close
			set rs_upd = nothing
			set rs_no_votes = nothing
			set rs_expir = nothing
			set rs_ip_blokck = nothing
		
		Response.Redirect(Request.ServerVariables("URL"))
	
	end if

end sub

Response.Write "</div></div></div></div>"
%>
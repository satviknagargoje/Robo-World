

      <%
	dim constr,rs,con
	
	constr ="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Inetpub\wwwroot\web-roboworld\roboworld.mdb;Persist Security Info=False"

	
	dim name,email,tel,fax,comment1,other,messageType,briefComment
	
	messageType= Request.form("messageType")
	comment1=Request.form("comment1")
	other=Request.form("other")
	briefComment=Request.form("briefComment")
	name=Request.form("name")
	email=Request.form("email")
	tel=Request.form("tel")
	fax=Request.form("fax")
	
	set con = server.CreateObject("ADODB.connection")
	con.open constr
	set rs = Server.CreateObject ("ADODB.Recordset")
	rs.open "Insert into feedback values('"&messageType&"','"&comment1&"','"&other&"','"&briefComment&"','"&name&"','"&email&"','"&tel&"','"&fax&"')",con
	set rs=nothing
	set con=nothing
	Response.Redirect "feedback.asp?name=" & name & "&email=" & email & "&tel=" & tel & "&fax=" & fax
      %>
  
 
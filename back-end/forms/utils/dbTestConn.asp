
<!--#include file="../utils/db.asp"-->

<%

function testConn(connectionStr)
  response.write(Request.ServerVariables("SERVER_NAME") & "<br><br>")
  response.write("Testing: " & connectionStr & "<br><br>")
  response.write("Starting: " & now & "<br>")
  set conn = Server.CreateObject("ADODB.Connection")
  conn.connectionString = connectionStr
    conn.open
		if failed then 
		  response.write("Could not connect to database.<br><br>" & errorMsg)
		else
      response.write("Finished OK: " & now)
		end if
  conn.close
  set conn = nothing
end function

dim s
s = request("connStr")
if len(s) > 0 then
  if s = "1" then
    testConn(connStr)
	else
    testConn(s)
	end if
else
  response.write("BAD_REQUEST")
end if

%>

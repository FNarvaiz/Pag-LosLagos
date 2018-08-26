<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001" %>
<% option explicit %>

<!--#include file="utils/db.asp"-->
<!--#include file="utils/usersDB.asp"-->
<!--#include file="formsData.asp"-->
<!--#include file="formsUtils.asp"-->
<!--#include file="app/appForms.asp"-->
<!--#include file="app/appProcesses.asp"-->

<!--#include file="utils/clsUpload.asp"-->

<%
dim objUpload: set objUpload = new clsUpload

sessionId = objUpload("sessionId").value
verb = objUpload("verb").value
if getLoggedUsrData then
  if not handleAppProcess then 
    response.write("BAD_REQUEST (" & verb & ")")
  end if
else
  response.write("NO_SESSION")
end if

set objUpload = nothing

%>


<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001" %>
<% option explicit %>
<%
response.expires=-1
response.charSet = "utf-8"
session.codePage = 65001

function zeroPad(value, length)
  if isNumeric(value) then
    zeroPad = String(length - len(value), "0") & value
  else
    zeroPad = value
  end if
end function

function fileLastMod(filename)
  dim fso, f, d
  set fso = server.createObject("Scripting.FileSystemObject")
  set f = fso.getFile(server.mapPath(filename))
  d = f.DateLastModified
  set f = nothing
  set fso = nothing
  fileLastMod = year(d) & zeroPad(month(d), 2) & zeroPad(day(d), 2)  & "-" & zeroPad(hour(d), 2) & zeroPad(minute(d), 2) & zeroPad(second(d), 2) 
end function

function includeStyleSheet(filename)
  %>
  <link href="<%= filename & "?version=" & fileLastMod(filename) %>" rel="stylesheet" type="text/css">
  <%
end function

function includeSript(filename)
  %>
  <script type="text/javascript" src="<%= filename & "?version=" & fileLastMod(filename) %>"></script>
  <%
end function

%>
<!--#include file="credits.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html lang="es">
<head>
<title>Barrio de los Lagos - Administración del sistema</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta http-equiv="Content-Script-Type" content="text/javascript">
<%
includeStyleSheet("forms/app/app.css")
includeStyleSheet("forms/app/basics.css")

includeStyleSheet("forms/forms.css")
includeStyleSheet("forms/formsGridView.css")
includeStyleSheet("forms/formsRecordView.css")
includeStyleSheet("forms/formsQueryOptions.css")
includeStyleSheet("forms/formsDialogs.css")
includeStyleSheet("forms/app/appForms.css")
includeStyleSheet("forms/datePicker.css")

includeSript("forms/app/app.js")
includeSript("forms/app/ajax.js")
includeSript("forms/app/elemState.js")

includeSript("forms/utils.js")
includeSript("forms/forms.js")
includeSript("forms/session.js")
includeSript("forms/datePicker.js")

'includeSript("../fckeditor/fckeditor.js")

%>
</head>
<body leftmargin="0" topmargin="0" onload="init();" onunload="done();" onresize="bodyResized()" onclick="datePickerHide();"></body>
</html>

<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001" %>
<% option explicit %>
<!--
  ***************************************************************

          DESARROLLO DE SOFTWARE
          Claudio Lago

          (c) 2007-<%= year(date) %> GBD - www.gbd.com.ar

  ***************************************************************
//-->
<!--#include file="front-end/forms/datePicker.asp"-->
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
  if fso.fileExists(server.mapPath(filename)) then
    set f = fso.getFile(server.mapPath(filename))
    d = f.DateLastModified
    fileLastMod = year(d) & zeroPad(month(d), 2) & zeroPad(day(d), 2)  & "-" & zeroPad(hour(d), 2) & zeroPad(minute(d), 2) & zeroPad(second(d), 2) 
  else
    fileLastMod = "FALTANTE"
  end if
  set f = nothing
  set fso = nothing
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
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>Vecinos de los Sauces - Sistema Brick</title>
<meta http-equiv="Content-Type" content="text/html;">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<meta http-equiv="Content-Script-Type" content="text/javascript">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<link rel="icon" type="image/png" href="contenidos/iconos/favicon-32x32.png" sizes="32x32" />
<link rel="icon" type="image/png" href="contenidos/iconos/favicon-16x16.png" sizes="16x16" />
<link rel="shortcut icon" href="contenidos/iconos/favicon.ico">
<%
includeStyleSheet("front-end/forms/forms.css")
includeStyleSheet("front-end/forms/formsDialogs.css")
includeStyleSheet("front-end/forms/formsGridView.css")
includeStyleSheet("front-end/forms/formsRecordView.css")
includeStyleSheet("front-end/forms/app/appForms.css")

includeSript("front-end/forms/forms.js")

includeStyleSheet("userfiles/estilos.css")
includeStyleSheet("front-end/main.css")
includeStyleSheet("front-end/timeLine.css")
includeStyleSheet("front-end/forms/datePicker.css")

includeSript("front-end/main.js")
includeSript("front-end/ajax.js")
includeSript("front-end/timeLine.js")
includeSript("front-end/misc.js")
includeSript("front-end/forms/datePicker.js")
includeSript("front-end/forms/utils.js")
includeStyleSheet("front-end/mediaQuery.css")


%>
<!--#include file="front-end/utils/googleAnalytics.asp"-->
</head>
<body leftmargin="0" topmargin="0" onload="init()" onunload="done()" onresize="bodyResized()" onclick="datePickerHide();">
  <div id="main">
    <div id="mainPanel">
      <div id="header"></div>
      <div id="mainMenu"></div>
      <div id="dynPanel"></div>
      <div id="dynPanelOverlay"></div>
      <div id="dataViewer">
        <div id="dataViewerPanelBg">
          <div id="dataViewerHideBtn" onclick="hideDataViewer()">X</div>
        </div>
        <div id="dataViewerPanel"></div>
      </div>
      <div id="imgViewer">
        <div id="imgViewerPanel"><img id="imgViewerImage" onload="this.style.visibility = 'visible';"></div>
        <div id="imgViewerHideBtn" onclick="hideImgViewer()">X</div>
      </div>
      <div id="footer">
        <div id="userMenu"></div>
        <div id="copyright">Sistema Brick © 2007-<%= year(date) %> GBD</div>
        <div id="gbdLogo" onclick="window.open('http://www.gbd.com.ar');"></div>
      </div>
    </div>
    <div id="dialogbackgnd"></div>
    <div id="dialogsPanel"></div>
    <div id="loadingSignal"></div>
  </div>
  <% renderDatePicker %>
</body>
</html>

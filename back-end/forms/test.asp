<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001" %>

<!--#include file="formsUtils.asp"-->
<!--#include file="utils/db.asp"-->

<%
response.expires=-1
response.charSet = "utf-8"
session.codePage = 65001
response.Buffer=true

sendCDOMail "Yo", "claudiolago@yahoo.com.ar", "test", "mensaje", "Vecinos", "info@vecinosloslagos.com.ar"
response.end


dbConnect
dbGetData("SELECT NOMBRE, CONTENIDO_TEXTO, ARCHIVO_FILENAME, ARCHIVO_CONTENTTYPE, ARCHIVO_BINARYDATA FROM LINEA_TIEMPO WHERE ID=947")
SendCDOMail "info@vecinosdelgolf.com.ar", "claudiolago@yahoo.com.ar", rs("NOMBRE").value, rs("CONTENIDO_TEXTO").value, "claudiomlago@yahoo.com.ar", _
  rs("ARCHIVO_FILENAME").value, rs("ARCHIVO_CONTENTTYPE").value, rs("ARCHIVO_BINARYDATA").value
dbDisconnect

%>

OK!!!


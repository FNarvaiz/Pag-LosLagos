<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001" %>
<% option explicit %>

<% dim connstr %>
<!--#INCLUDE FILE="clsUpload.asp"-->
<!--#INCLUDE FILE="dbConst.asp"-->
<%

function trySetRSFieldValue(fieldName, data)
  dim b: b = false
  dim f
  for each f in rs.fields
    if f.name = fieldName then
      f.value = data
      b = true
      exit for
    end if
  next
end function

function getFilter(keyFieldNames, keyFieldValues, recordId)
  dim keyFields, keyValues, str, i
  keyFields = split(keyFieldNames, ",")
  keyValues = split(keyFieldValues, ",")
  str = "ID=" & recordId
  for i = 0 to uBound(keyFields)
    str = str & " AND " & keyFields(i) & "=" & keyValues(i)
  next
	getFilter = str
end function

function updateBinaryData
  dim dbFieldBaseName: dbFieldBaseName = objUpload("dbFieldBaseName").value
  dim filename: filename = objUpload("uploadFormFileField").fileName
  if len(objUpload("uploadFormFileField").fileName) > 0 then
    response.write("Preparando para GRABAR el archivo: " & filename & "<br>")
    trySetRSFieldValue dbFieldBaseName & "_FILENAME", objUpload("uploadFormFileField").fileName
    trySetRSFieldValue dbFieldBaseName & "_FILEEXT", objUpload("uploadFormFileField").fileExt
    trySetRSFieldValue dbFieldBaseName & "_FILESIZE", objUpload("uploadFormFileField").length
    rs(dbFieldBaseName & "_CONTENTTYPE") = objUpload("uploadFormFileField").contentType
    rs(dbFieldBaseName & "_BINARYDATA").AppendChunk objUpload("uploadFormFileField").BLOB & ChrB(0)
  else
    response.write("Preparando para ELIMINAR el archivo: " & filename & "<br>")
    trySetRSFieldValue dbFieldBaseName & "_FILENAME", null
    trySetRSFieldValue dbFieldBaseName & "_FILEEXT", null
    trySetRSFieldValue dbFieldBaseName & "_FILESIZE", null
    rs(dbFieldBaseName & "_CONTENTTYPE") = null
    rs(dbFieldBaseName & "_BINARYDATA") = null
  end if
  response.write("Actualizando... ")
  rs.update
  if err.number = 0 then
    response.write("OK<br>")
  else
    response.write("Error: " & err.number & "<br>")
    err.clear
  end if
end function

on error resume next
dim objUpload: set objUpload = New clsUpload
dim conn: set conn = Server.CreateObject("ADODB.Connection")
conn.connectionString = connStr
conn.commandTimeout = 0
conn.open
dim sql: sql = " FROM " & objUpload("tableName").value & " WHERE " & _
  getFilter(objUpload("keyFields").value, objUpload("keyValues").value, objUpload("recordId").value)
dim rs: set rs = conn.Execute("SELECT count(*)" & sql)
dim count: count = rs(0)
rs.close
sql = "SELECT *" & sql
set rs = nothing
set rs = Server.CreateObject("ADODB.Recordset")
rs.open sql, conn, 2, 2, 1
select case count
  case 0: 
    response.write("ERROR: NO EXISTE EL REGISTRO<br>")
  case 1: 
    updateBinaryData
  case else
    response.write("ERROR: EXISTE MÁS DE 1 REGISTRO<br>")
end select
'response.write("keyFields=" & objUpload("keyFields").value & " - keyValues=" & objUpload("keyValues").value & _
'  " - recordId=" & objUpload("recordId").value & "<br>")

rs.close
set rs = nothing
conn.close
set conn = nothing
set objUpload = nothing

%>


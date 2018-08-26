
<!--#INCLUDE FILE="dbConst.asp"-->
<!--#INCLUDE FILE="JSON.asp"-->

<%
response.expires=-1
response.charSet = "utf-8"
session.codePage = 65001
response.Buffer=true
'On Error resume Next

dim connStr, conn, rs
set conn = nothing
set rs = nothing
if isEmpty(session("lastSQLCommand")) then session("lastSQLCommand") = ""

dim dbLogging: dbLogging = false

const sqlMsgPrefix = "[SQL Server]"

dim errorMsg
errorMsg = ""
function failed
  if err.number <> 0 then
    dim i
    i = inStr(err.description, sqlMsgPrefix)
    if i > 0 then 
      i = i + len(sqlMsgPrefix) - 1
      errorMsg = right(err.description, len(err.description) - i)
    else
      errorMsg = err.description
    end if
    err.clear
		failed = true
		if dbLogging then dbLog("FAILED: " & errorMsg)
	else
	  failed = false
  end if
end function

function dbConnect
  dim b: b = conn is nothing
  if b then
    if dbLogging then dbLog("CONNECT")
    set conn = Server.CreateObject("ADODB.Connection")
    conn.connectionString = connStr
    conn.commandTimeout = 0
    conn.open
    conn.Execute("SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED")
  end if
  dbConnect = b
end function

function dbDisconnect
  if dbLogging then dbLog("DISCONNECT")
  'dbGetData("SELECT @@TRANCOUNT")
  'if rs(0) > 0 then
  '  dbLog("OPEN TRANSACTION FOUND ON CONNECTION CLOSE: formId=" & formId & "; verb=" & verb & ", usrId=" & usrId) ' & "; keyValues=" & keyValues & "; recordId=" & recordId)
  '  dbExecute("ROLLBACK TRANSACTION")
  'end if
  conn.close
  set conn = nothing
end function

function dbGetData(sqlStr)
  if dbLogging then dbLog(sqlStr)
  session("lastSQLCommand") = sqlStr
  'set rs = conn.Execute(sqlStr)
  set rs = server.createObject("ADODB.recordset")
  rs.cursorLocation = 3 ' adUseClient
	rs.open sqlStr, conn, 0, 1, 1 ' adOpenStatic, adOpenForwardOnly, adCmdText
  set rs.activeConnection = nothing
  dbGetData = not rs.EOF
end function

function dbGetDataAux(sqlStr)
  'dbLog(sqlStr)
  session("lastSQLCommand") = sqlStr
  dim auxConn
  set auxConn = Server.CreateObject("ADODB.Connection")
  auxConn.connectionString = connStr
  auxConn.commandTimeout = 0
  auxConn.open
  auxConn.Execute("SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED")
  set rs = auxConn.Execute(sqlStr)
  dbGetDataAux = not rs.EOF
end function

function dbGetTable(tableName)' searchExpr
  set rs = Server.CreateObject("ADODB.Recordset")
  rs.open tableName, conn, 2, 2
'  if len(searchExpr) > 0 then rs.filter = chr(34) & searchExpr & chr(34)
  dbGetTable = not rs.EOF
end function

function dbReleaseData
  rs.close
	set rs = nothing
end function

function dbExecute(sqlStr)
  dim b: b = dbConnect
  if dbLogging then dbLog(sqlStr)
  session("lastSQLCommand") = sqlStr
	conn.Execute sqlStr, , 128 ' adExecuteNoRecords
	if b then dbDisconnect
end function

function dbLog(data)
  dim connLog
  set connLog = Server.CreateObject("ADODB.Connection")
  connLog.connectionString = connStr
  connLog.commandTimeout = 0
  connLog.open
  if isNull(data) then
    connLog.Execute("INSERT INTO DEBUG (DATA) VALUES ('DATA IS NULL')")
  else
    connLog.Execute("INSERT INTO DEBUG (DATA) VALUES (" & chr(39) & replace(data, "'", "''") & chr(39) & ")")
  end if
  connLog.close
  set connLog = nothing
end function

public Const dbSearchModeAll = 1
public Const dbSearchModeCurrent = 2

function getSearchExpr(searchMode, keyFieldNames, keyFieldValues, recordId)
  dim keyFields, keyValues, str, i, j, a
  keyFields = strToArray(keyFieldNames, ",")
  keyValues = strToArray(keyFieldValues, ",")
	str = ""
	select case searchMode
	  case dbSearchModeCurrent:
  	  str = "ID=" & recordId
    	for i = 0 to uBound(keyFields)
        str = str & " AND ("
        a = strToArray(keyFields(i), ";")
        for j = 0 to uBound(a)
          str = str & a(j) & "=" & keyValues(i)
          if j < uBound(a) then str = str & " OR "
        next
        str = str & ")"
    	next
		case dbSearchModeAll:
    	for i = 0 to uBound(keyFields)
    	  if i > 0 then str = str & " AND "
        str = str & "("
        a = strToArray(keyFields(i), ";")
        for j = 0 to uBound(a)
          str = str & a(j) & "=" & keyValues(i)
          if j < uBound(a) then str = str & " OR "
        next
        str = str & ")"
    	next
      if len(searchCondition) > 0 then ' searchCondition is a global var (see forms.asp)
        if len(str) > 0 then 
          str = str & " AND "
        end if
        str = str & searchCondition
      end if
	end select
	getSearchExpr = str
end function

function dbGetBinaryData(tableName, keyFieldNames, keyFieldValues, recordId, useFilename)
  dim baseFieldName: baseFieldName = getStringParam("dbFieldBaseName", 20)
  dim fileNameField: fileNameField = baseFieldName & "_FILENAME"
	dim contentTypeFieldName: contentTypeFieldName = baseFieldName & "_CONTENTTYPE"
	dim binaryDataFieldName: binaryDataFieldName = baseFieldName & "_BINARYDATA"
  if useFileName then
  	getRecord fileNameField & ", " & contentTypeFieldName & "," & binaryDataFieldName, tableName, keyFieldNames, keyFieldValues, recordId
  else
  	getRecord contentTypeFieldName & "," & binaryDataFieldName, tableName, keyFieldNames, keyFieldValues, recordId
  end if
	if rs.EOF then
	  response.write("Record not found.")
	else
    if isNull(rs(binaryDataFieldName)) then
      response.write("Null data.")
    else
      if useFilename then response.addHeader "Content-Disposition", "attachment; filename=" & dQuotes(rs(fileNameField))
  		response.contentType = rs(contentTypeFieldName)
  		response.binaryWrite rs(binaryDataFieldName)
    end if
	end if
  dbReleaseData
end function

function getRecord(fieldNames, tableName, keyFieldNames, keyFieldValues, recordId)
  getRecord = dbGetData("SELECT " & fieldNames & " FROM " & tableName & " WHERE " & _
	  getSearchExpr(dbSearchModeCurrent, keyFieldNames, keyFieldValues, recordId))
end function

function getAllRecords(tablename, keyFieldNames, keyFieldValues)
  getAllRecords = dbGetData("SELECT * FROM " & tableName & " WHERE " & getSearchExpr(dbSearchModeAll, keyFieldNames, keyFieldValues, null))
end function

function dbDelete(tableName, keyFieldNames, keyFieldValues, recordId)
  if usrProfile <> usrProfileIT then on error resume next
  if dbLogging then dbLog(tableName & " - delete - #" & recordId)
  dbExecute("DELETE FROM " & tableName & " WHERE " & getSearchExpr(dbSearchModeCurrent, keyFieldNames, keyFieldValues, recordId))
	if failed then
    dbDelete = false
    JSONAddOpFailed
    JSONAddMessage "No se pudo eliminar el elemento seleccionado.\n\nDetalles:\n\n" & errorMsg
	else
    dbDelete = true
    JSONAddOpOK
	end if
end function

function isDBFieldReadOnly(fieldsReadOnly, fieldIndex)
  if isNull(fieldsReadOnly) then
	  isDBFieldReadOnly = false
	else
	  isDBFieldReadOnly = fieldsReadOnly(fieldIndex)
  end if
end function

function dbGetLastId(tableName, keyFieldNames, keyFieldValues)
  dim searchAllExpr
  dim b: b = dbConnect
  searchAllExpr = getSearchExpr(dbSearchModeAll, keyFieldNames, keyFieldValues, null)
	if len(searchAllExpr) > 0 then searchAllExpr = " WHERE " & searchAllExpr
  dbGetData("SELECT MAX(ID) AS ID FROM " & tableName & searchAllExpr)
	if isNull(rs("ID")) then
	  dbGetLastId = 0
	else
    dbGetLastId = rs("ID")
	end if
	dbReleaseData
  if b then dbDisconnect
end function

function getIdValue(fieldNames, fieldValues)
  dim i
  for i = 0 to uBound(fieldNames)
	  if fieldNames(i) = "ID" then
		  getIdValue = fieldValues(i)
			exit for
		end if
	next
end function

function getValueFromLastRecordCreatedByUsr(fieldName)
  dim b: b = dbConnect
  if dbGetData("SELECT " & fieldName & " AS FIELDVALUE FROM " & formTable & _
    " WHERE ID=(SELECT MAX(ID) FROM " & formTable & " WHERE " & dbAuditUsrIdField & "=" & usrId & ")") then
    getValueFromLastRecordCreatedByUsr = rs("FIELDVALUE")
  else
    getValueFromLastRecordCreatedByUsr = null
  end if
	dbReleaseData
  if b then dbDisconnect
end function

const dbDecimalPrefix = "[DECIMAL]"
const dbMoneyPrefix = "[MONEY]"
const dbDatePrefix = "[DATE]"
const dbTimePrefix = "[TIME]"
const dbDateSeparator = "/"
const dbTimeSeparator = ":"

function isDecimalValue(val)
  isDecimalValue = (left(val, len(dbDecimalPrefix)) = dbDecimalPrefix)
end function

function isMoneyValue(val)
  isMoneyValue = (left(val, len(dbMoneyPrefix)) = dbMoneyPrefix)
end function

function isDateValue(val)
  isDateValue = (left(val, len(dbDatePrefix)) = dbDatePrefix)
end function

function isTimeValue(val)
  isTimeValue = (left(val, len(dbTimePrefix)) = dbTimePrefix)
end function

function stripDecimalPrefix(str)
  stripDecimalPrefix = right(str, len(str) - len(dbDecimalPrefix))
end function

function stripMoneyPrefix(str)
  stripMoneyPrefix = right(str, len(str) - len(dbMoneyPrefix))
end function

function stripDatePrefix(str)
  stripDatePrefix = right(str, len(str) - len(dbDatePrefix))
end function

function stripTimePrefix(str)
  stripTimePrefix = right(str, len(str) - len(dbTimePrefix))
end function

function sqlTime(str)
  if str = dbTimePrefix or isNull(str) then
    sqlTime = "NULL"
  else
    sqlTime = ""
    dim timeParts: timeParts = split(stripTimePrefix(str), dbTimeSeparator)
    dim t
    if uBound(timeParts) >= 2 then
      t = timeSerial(timeParts(0), timeParts(1), timeParts(2))
    elseif uBound(timeParts) = 1 then
      t = timeSerial(timeParts(0), timeParts(1), 0)
    elseif isArray(timeParts) then
      if isNumeric(timeParts(0)) then
        t = timeSerial(timeParts(0), 0, 0)
      else
        err.raise vbObjectError + 1000, "sqlTime", "Formato de hora incorrecto. Ingrese hh:mm."
      end if
    else
      err.raise vbObjectError + 1000, "sqlTime", "Formato de hora incorrecto. Ingrese hh:mm."
    end if
    redim timeParts(2)
    timeParts(0) = hour(t)
    timeparts(1) = minute(t)
    timeParts(2) = second(t)
    sqlTime = sQuotes(timeParts(0) & dbTimeSeparator & timeParts(1) & dbTimeSeparator & timeParts(2))
  end if
end function

function sqlDate(str)
  if str = dbDatePrefix or isNull(str) then
    sqlDate = "NULL"
  else
    dim dateParts: dateParts = split(stripDatePrefix(str), dbDateSeparator)
    dim d
    if uBound(dateParts) >= 2 then
      d = dateSerial(dateParts(2), dateParts(1), dateParts(0))
    elseif uBound(dateParts) = 1 then
      d = dateSerial(year(date), dateParts(1), dateParts(0))
    elseif isArray(dateParts) then
      if isNumeric(dateParts(0)) then
        d = dateSerial(year(date), month(date), dateParts(0))
      else
        err.raise vbObjectError + 1000, "sqlDate", dateParts(0) & " tiene un formato de fecha incorrecto. Ingrese dd/mm/aaaa."
      end if
    else
      err.raise vbObjectError + 1000, "sqlDate", dateParts(0) & " tiene un formato de fecha incorrecto. Ingrese dd/mm/aaaa."
    end if
    redim dateParts(2)
    dateParts(0) = year(d)
    dateParts(1) = month(d)
    dateParts(2) = day(d)
    sqlDate = sQuotes(dateParts(0) & zeroPad(dateParts(1), 2) & zeroPad(dateParts(2), 2))
  end if
end function

function sqlDecimal(str)
  if str = dbDecimalPrefix then
    sqlDecimal = "NULL"
  else
    sqlDecimal = systemToSqlDecimal(stripDecimalPrefix(str))
  end if
end function

function sqlMoney(str)
  if str = dbMoneyPrefix then
    sqlMoney = "NULL"
  else
    sqlMoney = "CAST(" & sQuotes(systemToSqlDecimal(stripMoneyPrefix(str))) & " AS MONEY)"
  end if
end function

function sqlInt(value)
  if isEmpty(value) or not isNumeric(value) then
    sqlInt = "NULL"
  else
    sqlInt = cLng(value)
  end if
end function

function sqlValue(fieldValue)
  if isNull(fieldValue) then
    sqlValue = "NULL"
  elseif len(fieldValue) = 0 then
    sqlValue = "NULL"
  elseif isDecimalValue(fieldValue) then
    sqlValue = sqlDecimal(fieldValue)
  elseif isMoneyValue(fieldValue) then
    sqlValue = sqlMoney(fieldValue)
  elseif isDateValue(fieldValue) then
    sqlValue = sqlDate(fieldValue)
  elseif isTimeValue(fieldValue) then
    sqlValue = sqlTime(fieldValue)
  else
    dim v: v = systemToSqlDecimal(fieldValue)
    if isNumeric(v) then
      if v = int(v) then
        sqlValue = v
      else
        sqlValue = sQuotes(v)
      end if
    else
      dim t: t = split(fieldValue, " ")
      if uBound(t) = 1 then
        if isDate(t(0)) and isDate(t(1)) then
          sqlValue = sQuotes(year(t(0)) & zeroPad(month(t(0)), 2) & zeroPad(day(t(0)), 2) & " " & t(1))
        else
          sqlValue = "N" & sQuotes(fieldValue)
        end if
      else
        sqlValue = "N" & sQuotes(fieldValue)
      end if
    end if
  end if
end function

function systemToSqlDecimal(str)
  if isNull(str) then
    systemToSqlDecimal = "NULL"
  else
    systemToSqlDecimal = replace(replace(str, ".", ""), ",", ".")
  end if
end function

function systemToSqlDate(dateVal)
  if isNull(dateVal) then
    systemToSqlDate = "NULL"
  else
    systemToSqlDate = sQuotes(year(dateVal) & zeroPad(month(dateVal), 2) & zeroPad(day(dateVal), 2))
  end if
end function

function dbUpdate(tableName, keyFieldNames, keyFieldValues, fieldNames, fieldChanged, fieldValues, fieldsReadOnly, recordId, _
  idFieldIsIdentity, useAuditData)
  if usrProfile <> usrProfileIT then on error resume next
  dbUpdate = -1
  dim s, t, i, sql, idValue: idValue = -1
	if cDbl(recordId) >= 0 then
    if useAuditData then 
      s = dbAuditUsrIdField & "=" & usrId & ", " & dbAuditDate & "=GETDATE() "
    else
      s = ""
    end if
		for i = 0 to UBound(fieldNames)
		  if not isDBFieldReadOnly(fieldsReadOnly, i) and fieldChanged(i) then
  			if len(s) > 0 then s = s & ", "
        s = s & fieldNames(i) & "=" & sqlValue(fieldValues(i))
			end if
		next
    if err.number = 0 then
      if len(s) > 0 then
        if idFieldIsIdentity then
          sql = "UPDATE " & tableName & " SET " & s & " WHERE ID=" & recordId
        else
          sql = "UPDATE " & tableName & " SET " & s & " WHERE " & getSearchExpr(dbSearchModeCurrent, keyFieldNames, keyFieldValues, recordId)
        end if
        dbExecute(sql)
      end if
      idValue = cDbl(getIdValue(fieldNames, fieldValues))
    end if
  else
    if dbLogging then dbLog(tableName & " - insert")
    s = keyFieldNames
    t = keyFieldValues
    if useAuditData then 
      if len(s) > 0 then s = ", " & s : t = ", " & t
      s = dbAuditUsrIdField & ", " & dbAuditDate & s
      t = usrId & ", " & "GETDATE() " & t
    end if
		for i = 0 to UBound(fieldNames)
		  if not isDBFieldReadOnly(fieldsReadOnly, i) then
  			if len(s) > 0 then s = s & ", " : t = t & ", "
  			s = s & fieldNames(i)
        t = t & sqlValue(fieldValues(i))
			end if
		next
	  sql = "INSERT INTO " & tableName & " (" & s & ") VALUES (" & t & ")"
		if idFieldIsIdentity then
      dbGetData("SET NOCOUNT ON " & sql & " SELECT ID=SCOPE_IDENTITY() SET NOCOUNT OFF")
      if err.number = 0 then
        idValue = cDbl(rs("ID"))
        dbReleaseData
      end if
		else
      if err.number = 0 then
        dbExecute(sql)
        idValue = cDbl(getIdValue(fieldNames, fieldValues))
      end if
		end if
	end if
  if failed then
    JSONAddOpFailed
    JSONAddMessage "No se pudo guardar el elemento.\n\nDetalles:\n\n" & errorMsg
  else
    JSONAddOpOK
    JSONAdd "recordId", idValue
    dbUpdate = idValue
  end if
end function

public const dbMoveRecordUp = 1
public const dbMoveRecordDown = 2

function dbMoveRecord(direction, tableName, keyFieldNames, keyFieldValues, recordId)
  if isNull(direction) then
    response.write("BAD_REQUEST")
    exit function
  end if
  dim maxId, minId, auxId, doIt
  dim b: b = dbConnect
  searchAllExpr = getSearchExpr(dbSearchModeAll, keyFieldNames, keyFieldValues, null)
	if searchAllExpr <> "" then searchAllExpr = " WHERE " & searchAllExpr
  dbGetData("SELECT MAX(ID) AS ID FROM " & tableName & searchAllExpr)
  maxId = rs("ID")
	dbReleaseData
	doIt = false
	select case direction
	  case dbMoveRecordUp:
      dbGetData("SELECT MIN(ID) AS ID FROM " & tableName & searchAllExpr)
      minId = rs("ID")
    	dbReleaseData
      if recordId > minId then
    		if searchAllExpr <> "" then searchAllExpr = searchAllExpr & " AND "
        dbGetData("SELECT MAX(ID) AS ID FROM " & tableName & searchAllExpr & " ID<" & recordId)
      	auxId = rs("ID")
    		dbReleaseData
				doIt = true
			end if
		case dbMoveRecordDown:
      if maxId > recordId then
    		if searchAllExpr <> "" then searchAllExpr = searchAllExpr & " AND "
        dbGetData("SELECT MIN(ID) AS ID FROM " & tableName & searchAllExpr & " ID>" & recordId)
      	auxId = rs("ID")
    		dbReleaseData
				doIt = true
			end if
		case else: auxId = recordId
	end select
  if doIt then
		dbUpdate tableName, keyFieldNames, keyFieldValues, array("ID"), array(maxId + 1), recordId, false, false
		dbUpdate tableName, keyFieldNames, keyFieldValues, array("ID"), array(recordId), auxId, false, false
		dbUpdate tableName, keyFieldNames, keyFieldValues, array("ID"), array(auxId), maxId + 1, false, false
  end if
	if b then dbDisconnect
	if failed then
    dbMoveRecord = false
    JSONAddOpFailed
    JSONAddMessage "No se pudo mover el elemento seleccionado.\n\nDetalles:\n\n" & errorMsg
	else
    dbMoveRecord = true
    JSONAddOpOK
    JSONAdd "recordId", auxId
	end if
  JSONSend
end function


%>
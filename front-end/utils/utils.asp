<%

setLocale(11274)

const uiDecimalSeparator = ","
dim sysDecimalSeparator
if isNumeric("1,1") then
  sysDecimalSeparator = ","
else
  sysDecimalSeparator = "."
end if
const notAvailableFieldValue = "N/D"
const notApplicableFieldValue = "N/A"

function sQuotes(str)
  sQuotes = chr(39) & str & chr(39)
end function

function dQuotes(str)
  dQuotes = chr(34) & str & chr(34)
end function

function escapeQuotes(str)
  escapeQuotes = replace(str, chr(39), chr(39) & chr(39))
end function

function addSearchExpr(searchExpr)
  if len(searchExpr) = 0 then exit function
  if len(searchCondition) > 0 then
    searchCondition = searchCondition & " AND (" & searchExpr & ")"
  else
    searchCondition = "(" & searchExpr & ")"
  end if
end function

function addSearchValues(field, values)
  if len(values) = 0 then exit function
  dim i, a: a = strToArray(values, " ")
  for i = 0 to uBound(a)
    if len(a(i)) > 0 then
      a(i) = replace(replace(replace(replace(a(i), "[", "[[]"), "_", "[_]"), "?", "_"), ".", "")
      if left(a(i), 1) = "-" then
        a(i) = right(a(i), len(a(i)) - 1)
        addSearchExpr("REPLACE(" & field & ", '.', '') NOT LIKE '%' + " & sQuotes(a(i)) & " + '%'")
      else
        addSearchExpr("REPLACE(" & field & ", '.', '') LIKE '%' + " & sQuotes(a(i)) & " + '%'")
      end if
    end if
  next
end function

function addSearchList(field, list)
  if len(list) = 0 then exit function
  dim i, a: a = strToArray(replace(list, vbCr, " "), " ")
  dim s: s = ""
  for i = 0 to uBound(a)
    if len(s) > 0 then s = s & " OR "
    if len(a(i)) > 0 then s = s & "(" & field & " LIKE '%" & replace(replace(replace(a(i), "[", "[[]"), "_", "[_]"), "?", "_") & "%'" & ")"
  next
  if len(s) > 0 then addSearchExpr(s)
end function

function formatDecimal(str, decimals)
  if isNull(str) then
    formatDecimal = ""
  else
    formatDecimal = formatNumber(str, decimals)
  end if
end function

function strToDecimal(str)
  strToDecimal = null
  if isNull(str) then exit function
  if inStr(str, uiDecimalSeparator) > 0 then
    str = replace(replace(str, ".", ""), uiDecimalSeparator, sysDecimalSeparator)
  else
    str = replace(str, ".", sysDecimalSeparator)
  end if
  if isNumeric(str) then strToDecimal = cDbl(str)
end function

function strToInteger(str)
  strToInteger = null
  if isNull(str) then exit function
  if inStr(str, uiDecimalSeparator) > 0 then
    str = replace(replace(str, ".", ""), uiDecimalSeparator, sysDecimalSeparator)
  else
    str = replace(str, ".", sysDecimalSeparator)
  end if
  if isNumeric(str) then strToInteger = cLng(str)
end function

function strToArray(str, delimiter)
  if isNull(str) or len(str) = 0 then
    strToArray = array()
  else
    strToArray = split(str, delimiter)
  end if
end function

function zeroPad(value, length)
  if isNumeric(value) then
    zeroPad = String(length - len(value), "0") & value
  else
    zeroPad = value
  end if
end function

function singleLine(str)
  if isNull(str) then
    singleLine = ""
  else
    singleLine = replace(str, " ", "&nbsp;")
  end if
end function

function cleanupInputParam(str)
  cleanupInputParam = replace(replace(replace(replace(replace(str, chr(34), ""), chr(39), ""), ";", ""), ":", ""), "--", "")
end function

function isValidParamName(paramName)
  isValidParamName = false
  if isEmpty(paramName) then exit function
  if isNull(paramName) then exit function
  if len(paramName) = 0 then exit function
  isValidParamName = true
end function

function getNumericParam(paramName)
  getNumericParam = null
  if not isValidParamName(paramName) then exit function
  dim i: i = request(paramName)
  if len(i) = 0 then exit function
  if isNumeric(i) then getNumericParam = strToDecimal(i)
end function

function getIntegerParam(paramName)
  getIntegerParam = null
  if not isValidParamName(paramName) then exit function
  dim i: i = request(paramName)
  if len(i) = 0 then exit function
  if isNumeric(i) then getIntegerParam = strToInteger(i)
end function

function getStringParam(paramName, maxLength)
  getStringParam = null
  if not isValidParamName(paramName) then exit function
  dim s: s = trim(request(paramName))
  if len(s) = 0 then exit function
  s = escapeQuotes(s)
  if maxLength > 0 then s = left(s, maxLength)
  getStringParam = s
end function

function getDateTimeParam(paramName)
  getDateTimeParam = null
  if not isValidParamName(paramName) then exit function
  dim d: d = request(paramName)
  if isDate(d) then getDateTimeParam = d
end function

function isValidCUIT(str)
  isValidCUIT = false
  str = replace(str, "-", "")
  if len(str) = 11 then exit function
  if not isNumeric(str) then exit function
  dim chars_1_2: chars_1_2 = str(0) & str(1)
  if chars_1_2 <> "20" and chars_1_2 <> "23" and chars_1_2 <> "24" and chars_1_2 <> "27" and chars_1_2 <> "30" and _
      chars_1_2 <> "33" and chars_1_2 <> "34" then exit function
  dim a: a = array(5, 4, 3, 2, 7, 6, 5, 4, 3, 2, 1)
  dim i, j: j = 0
  for i = 0 to uBound(a)
    j = j + asc(mid(str, i + 1, 1)) * a(i)
  next
  isValidCUIT = (j mod 11 = 0)
end function

function initTextStream(byRef stream)
  stream.Type = 2 'Text
  stream.Open
end function

function sendTextStream(byRef stream)
  stream.position = 0
  response.write(stream.readText)
end function

function nameCase(str)
  nameCase = null
  if isNull(str) then exit function
  if len(str) = 0 then exit function
  dim s, c, i
  dim delimiters: delimiters = " ,.-/_"
  i = 1
  do while i <= len(str)
    c = mid(str, i, 1)
    do while inStr(Delimiters, c) > 0 and i <= len(str)
      s = s + c
      i = i + 1
      c = mid(str, i, 1)
    loop
    if c = lCase(c) then
      s = s & uCase(c)
    else
      s = s & c
    end if
    i = i + 1
    do while i <= len(str)
      c = mid(str, i, 1)
      if inStr(Delimiters, c) > 0 then 
        exit do
      elseif c = uCase(c) then
        s = s & lCase(c)
      else
        s = s & c
      end if
      i = i + 1
    loop
  loop
  nameCase = s
end function

function sendMail2(name, email, subject, message)
  dim mailMsg
  set mailMsg = Server.CreateObject("Persits.MailSender")
  mailMsg.Host = "190.210.151.238"
  mailMsg.From = "info@vecinosloslagos.com.ar"
  mailMsg.FromName = "Vecinos de los Lagos"
  mailMsg.AddAddress email, name
  mailMsg.Subject = subject
  mailMsg.isHTML = true
  mailMsg.Body = name & ",<br><br>" & message & "<br><br>Este mensaje es una notificación automática. No lo responda, gracias."
  mailMsg.Send
  sendMail2 = err '0 = OK
  set mailMsg = nothing
end function


%>
<%

dim JSONResponse: JSONResponse = ""

const JSONOpResultField = "result"
const JSONOpOK = "OK"
const JSONOpFailed = "failed"
const JSONOpErrCodeField = "errCode"
const JSONOpResultMessageField = "message"

function JSONSend
  if len(JSONResponse) > 0 then 
    response.clear
    response.write("{" & JSONResponse & "}")
  end if
end function

function JSONAdd(name, var)
  if len(JSONResponse) > 0 then JSONResponse = JSONResponse & ","
  JSONResponse = JSONResponse & name & ":" & JSONValue(var)
end function

function JSONAddStr(name, str)
  if len(JSONResponse) > 0 then JSONResponse = JSONResponse & ","
  JSONResponse = JSONResponse & name & ":" & dQuotes(JSONescape(str))
end function

function JSONAddArray(name, arr, colNames)
  if len(JSONResponse) > 0 then JSONResponse = JSONResponse & ","
  JSONResponse = JSONResponse & name & ":" & JSONArray(arr, colNames)
end function

function JSONAddOpOK
  JSONAddStr JSONOpResultField, JSONOpOK
end function

function JSONAddOpFailed
  JSONAddStr JSONOpResultField, JSONOpFailed
end function

function JSONAddErrorCode(errCodeStr)
  JSONAddStr JSONOpErrCodeField, errCodeStr
end function

function JSONAddMessage(messageStr)
  JSONAddStr JSONOpResultMessageField, messageStr
end function

function JSONReportResult
  if failed then
    JSONAddOpFailed
  else
    JSONAddOpOK
  end if
end function

function JSONValue(value)
  if isNull(value) then
    JSONValue = "null"
  else
    dim vType
    vType = varType(value)
    if vType = 11 then 'boolean
      if value then JSONValue = "true" else JSONValue = "false"
    elseif vType = 2 or vType = 3 or vType = 17 or vType = 19 then 'int, long, byte
      JSONValue = cLng(value)
    elseif vType = 4 or vType = 5 or vType = 6 or vType = 14 then 'single, double, currency and type 14 (db)
      JSONValue = replace(cDbl(value), ",", ".")
    else
      JSONValue = dQuotes(JSONEscape(value))
    end if
  end if
end function

function JSONArray(byRef arr, colNames)
  JSONArray = "[ "
  if isArray(arr) then
    dim r, c
    for r = 0 to uBound(arr, 2)
      if r > 0 then JSONArray = JSONArray & ","
      JSONArray = JSONArray & "{ "
      for c = 0 to uBound(arr, 1)
        if c > 0 then JSONArray = JSONArray & ","
        JSONArray = JSONArray & colNames(c) & ":" & JSONValue(arr(c, r))
      next
      JSONArray = JSONArray & " }"
    next
  end if
  JSONArray = JSONArray & " ]"
end function

function JSONEscape(value)
  JSONEscape = ""
  if isEmpty(value) then exit function
  if isNull(value) then exit function
  dim str: str = cStr(value)
  dim i, c, wChr
  dim s: s = ""
  for i = 1 to len(str)
    c = mid(str, i, 1)
    wChr = ascW(c)
    if (wChr > &h00 and wChr < &h1F) or wChr = &h22 then
      c = "\u00" + zeroPad(hex(wChr), 2)
    elseif wChr >= &hC280 and wChr <= &hC2BF then
      c = "\u00" + zeroPad(hex(wChr - &hC200), 2)
    elseif wChr >= &hC380 and wChr <= &hC3BF then
      c = "\u00" + zeroPad(hex(wChr - &hC2C0), 2)
    end if
    s = s & c
  next
  JSONescape = s
end function

%>

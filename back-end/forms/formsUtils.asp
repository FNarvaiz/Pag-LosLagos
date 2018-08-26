<SCRIPT LANGUAGE=javaScript RUNAT=SERVER>
function functionExists(funcName)
{
  if (!funcName) return(false);
  var type = "";
  try
  {
    type = eval("typeof " + funcName);
  }
  catch (e) 
  {
    return(false);
  }
  return(type == "unknown" || type == "function");
}

function isValidEmail(emailAddress) {
  var re = /^(([^<>()[\]\\.,;:\s@"]+(\.[^<>()[\]\\.,;:\s@"]+)*)|(".+"))@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}])|(([a-zA-Z\-0-9]+\.)+[a-zA-Z]{2,}))$/;

  var sQtext = '[^\\x0d\\x22\\x5c\\x80-\\xff]';
  var sDtext = '[^\\x0d\\x5b-\\x5d\\x80-\\xff]';
  var sAtom = '[^\\x00-\\x20\\x22\\x28\\x29\\x2c\\x2e\\x3a-\\x3c\\x3e\\x40\\x5b-\\x5d\\x7f-\\xff]+';
  var sQuotedPair = '\\x5c[\\x00-\\x7f]';
  var sDomainLiteral = '\\x5b(' + sDtext + '|' + sQuotedPair + ')*\\x5d';
  var sQuotedString = '\\x22(' + sQtext + '|' + sQuotedPair + ')*\\x22';
  var sDomain_ref = sAtom;
  var sSubDomain = '(' + sDomain_ref + '|' + sDomainLiteral + ')';
  var sWord = '(' + sAtom + '|' + sQuotedString + ')';
  var sDomain = sSubDomain + '(\\x2e' + sSubDomain + ')*';
  var sLocalPart = sWord + '(\\x2e' + sWord + ')*';
  var sAddrSpec = sLocalPart + '\\x40' + sDomain; // complete RFC822 email address spec
  var sValidEmail = '^' + sAddrSpec + '$'; // as whole string
  var reValidEmail = new RegExp(sValidEmail);

  return(reValidEmail.test(emailAddress) && re.test(emailAddress));
}
</SCRIPT>

<%

function sQuotes(str)
  sQuotes = chr(39) & str & chr(39)
end function

function dQuotes(str)
  dQuotes = chr(34) & str & chr(34)
end function

function escapeQuotes(str)
  if isNull(str) then
    escapeQuotes = null
  else
    escapeQuotes = replace(str, chr(39), chr(39) & chr(39))
  end if
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

function sendPersitsMail(name, email, subject, message)
  on error resume next
  dim mailMsg
  set mailMsg = Server.CreateObject("Persits.MailSender")
  mailMsg.Host = "190.210.151.238"
  mailMsg.From = mainEmailAddress
  mailMsg.FromName = mailMsg.EncodeHeader(mainEmailName, "utf-8")
  mailMsg.AddAddress email,mailMsg.EncodeHeader(name, "utf-8")
  mailMsg.Subject = mailMsg.EncodeHeader(subject, "utf-8")
  mailMsg.isHTML = true
  mailMsg.Body = name & ",<br><br>" & message
  mailMsg.AddCustomHeader("Content-Transfer-Encoding: Quoted-Printable")
  mailMsg.AddCustomHeader("Content-Type: text/html; charset=" & dQuotes("UTF-8"))
'  mailMsg.CharSet = "UTF-8"
'  mailMsg.ContentTransferEncoding = "Quoted-Printable"
  mailMsg.Send
  sendPersitsMail = err '0 = OK
  set mailMsg = nothing
end function

function sendPersitsGroupMailPrepare(byRef mailMsg, subject, message, attachmentFilename, attachementData)
  mailMsg.resetAll
  mailMsg.Host = "190.210.151.238"
  mailMsg.From = mainEmailAddress
  mailMsg.FromName = mainEmailName 'mailMsg.EncodeHeader(mainEmailName, "utf-8")
  mailMsg.Subject = mailMsg.EncodeHeader(subject, "utf-8")
  mailMsg.isHTML = true
  mailMsg.Body = message
  if not isNull(attachmentFilename) and not isNull(attachementData) then
    mailMsg.AddAttachmentMem attachmentFilename, attachementData
  end if
  mailMsg.AddCustomHeader("Content-Transfer-Encoding: Quoted-Printable")
  mailMsg.AddCustomHeader("Content-Type: text/html; charset=" & dQuotes("UTF-8"))
'  mailMsg.CharSet = "UTF-8"
'  mailMsg.ContentTransferEncoding = "Quoted-Printable"
end function

function sendPersitsGroupMail(emailAddresses, occultEmailAddresses, subject, message, attachmentFilename, attachementData)
  on error resume next
  'emailAddresses = "clago@gbd.com.ar"
  'occultEmailAddresses = "claudiomlago@yahoo.com.ar"
  dim result: result = ""
  dim mailMsg
  set mailMsg = Server.CreateObject("Persits.MailSender")
  sendPersitsGroupMailPrepare mailMsg, subject, message, attachmentFilename, attachementData
  if len(emailAddresses) > 0 then mailMsg.AddAddress emailAddresses
  if len(occultEmailAddresses) > 0 then 
    dim a: a = split(occultEmailAddresses, ", ")
    dim i
    dim j: j = 0
    for i = 0 to uBound(a)
      if isValidEmail(a(i)) then 
        mailMsg.AddBcc a(i)
        j = j + 1
      else
        dbLog("Mail no válido: " & a(i))
      end if
      if j > 0 and ((i + 1) mod 10 = 0 or i = uBound(a)) then 
        j = 0
        mailMsg.Send
        if err.number <> 0 then 
          result = replace(err.description, vbCrLf, "")
          exit for
        end if
        sendPersitsGroupMailPrepare mailMsg, subject, message, attachmentFilename, attachementData
      end if
    next
  end if
  set mailMsg = nothing
  sendPersitsGroupMail = result '0 = OK
end function

function sendCDOMail(sName, sEmail, sSubject, sMessage, senderName, senderEmail)
  on error resume next
	dim cdoMail: set cdoMail = server.CreateObject("CDO.Message")
	dim cdoConf: set cdoConf = server.createObject("CDO.Configuration") 
	dim cdoFields: set cdoFields = cdoConf.fields
	with cdoFields
		.item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
		.item("http://schemas.microsoft.com/cdo/configuration/smtpserver")  = "localhost"
		.Item ("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
    .Item ("http://schemas.microsoft.com/cdo/configuration/sendusername") = "info@vecinosloslagos.com.ar"
    .Item ("http://schemas.microsoft.com/cdo/configuration/sendpassword") = PassMails
		.update
	end with
	with cdoMail
		set .configuration = cdoConf
		.bodyPart.Charset = "utf-8"
		.from = senderName & " <" & senderEmail & ">"
		.to = sName & " <" & sEmail & ">"
		.Bcc = senderName & " <" & senderEmail & ">"
		.subject = sSubject
		.HTMLBody = sMessage
		.send
	end with
	set cdoFields = nothing
	set cdoConf = nothing
	set cdoMail = nothing
End function

function validateEmailAddresses(addresses)
  dim invalidAddresses: invalidAddresses = ""
  dim a: a = split(addresses, ",")
  dim i
  for i = 0 to uBound(a)
    a(i) = trim(a(i))
    if not isValidEmail(a(i)) then
      if len(invalidAddresses) > 0 then invalidAddresses = invalidAddresses & ","
      invalidAddresses = invalidAddresses & a(i)
    end if
  next
  validateEmailAddresses = invalidAddresses
end function

function sendCDOGroupMail(emailAddresses, occultEmailAddresses, mailSubject, message, attachmentFilename, attachmentContentType, attachementData)
  on error resume next
  dim result: result = ""
	dim cdoMail: set cdoMail = server.CreateObject("CDO.Message")
	dim cdoConf: set cdoConf = server.createObject("CDO.Configuration") 
	dim cdoFields: set cdoFields = cdoConf.fields
	with cdoFields
		.item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
    .item("http://schemas.microsoft.com/cdo/configuration/smtpserver")  = "localhost"
    .Item ("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
    .Item ("http://schemas.microsoft.com/cdo/configuration/sendusername") = "info@vecinosloslagos.com.ar"
    .Item ("http://schemas.microsoft.com/cdo/configuration/sendpassword") = PassMails
		.update
	end with
	with cdoMail
		set .configuration = cdoConf
		.bodyPart.Charset = "utf-8"
		.from = mainEmailAddress
		.subject = mailSubject
		.textBody = message
	end with

	if not isNull(attachmentFilename) and not isNull(attachmentContentType) and not isNull(attachementData) then
    dim attach: set attach = cdoMail.Attachments.Add
    with attach
      .ContentMediaType = attachmentContentType
      .ContentTransferEncoding = "base64"
      .Fields("urn:schemas:mailheader:content-disposition").Value = "attachment;filename=" & dQuotes(attachmentFilename)
      .Fields.update
    end with
    dim stream: set stream = attach.GetDecodedContentStream
    stream.Write attachementData
    stream.Flush
  end if

  if len(emailAddresses) > 0 then cdoMail.to = replace(emailAddresses, ",", ";")
  if len(occultEmailAddresses) > 0 then 
    dim a: a = split(occultEmailAddresses, ",")
    dim i
    dim j: j = 0
    dim addressList: addressList = ""
    for i = 0 to uBound(a)
      a(i) = trim(a(i))
      if isValidEmail(a(i)) then 
        if len(addressList) > 0 then addressList = addressList & ";"
        addressList = addressList & a(i)
        j = j + 1
      else
        'dbLog("Mail no válido: " & a(i))
      end if
      if j > 0 and ((i + 1) mod ((uBound(a) + 1) \ 10) = 0 or i = uBound(a)) then 
        j = 0
        'dbLog(addressList)
        cdoMail.bcc = addressList
        cdoMail.send
        addressList = ""
        if err.number <> 0 then 
          result = replace(err.description, vbCrLf, "")
          exit for
        end if
      end if
    next
  end if
  set stream = nothing
	set cdoFields = nothing
	set cdoConf = nothing
	set cdoMail = nothing
  sendCDOGroupMail = result
end function



%>
<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001" %>
<% option explicit %>

<!--#include file="utils/db.asp"-->
<!--#include file="utils/usersDB.asp"-->
<!--#include file="formsData.asp"-->
<!--#include file="formsGridView.asp"-->
<!--#include file="formsRecordView.asp"-->
<!--#include file="formsQueryOptions.asp"-->
<!--#include file="formsDialogs.asp"-->
<!--#include file="formsUtils.asp"-->
<!--#include file="datePicker.asp"-->
<!--#include file="app/appForms.asp"-->

<%

function renderFormsContainer
	%>
	<div class="canvas formContainer <%= formContainerCssClass %>" id="<%= formId %>">
	<%
	if formTitle <> "" then
	  %>
    <div class="canvas formContainerTitle <%= formContainerTitleCssClass %>">
      <span id="formContainerTitleText"><%= formTitle %></span><span id="<%= formId %>SelectedKey"></span>
    </div>
	  <%
	end if
  if usrAccessAdminMaster then
    formInService = true
'  else
'    formInService = true ' Implementar !!
  end if
  if formInService then 
    if not isNull(forms) then
      %>
      <div class="hidden" id="<%= formId %>ChildFormIds"><%= childFormIds %></div>
      <div class="canvas formContainerMainPanel" id="<%= formId %>MainPanel">
        <%
        dim i
        for i = 0 to UBound(forms)
          %>
            <div class="canvas form <%= formCssClass %>" id="<%= forms(i) %>"></div>
          <%
        next
        %>
      </div>
      <%
    end if
  else
    %>
      <div class="hidden" id="<%= formId %>ChildFormIds">none</div>
      <div class="canvas form  <%= formCssClass %>">
        <table width="100%" height="100%"><tr><td align="center" valign="middle"><b>Este módulo se encuentra momentáneamente fuera de servicio</b></td></tr></table>
      </div>
    <%
  end if
  %>
	</div>
	<%
end function

function renderDBForm
  if formTitle <> "" then
  	%>
    <div class="canvas formTitle <%= formTitleCssClass %>" id="<%= formId %>FormTitle"><%= formTitle %></div>
    <%
	end if
	%>
	<div class="hidden" id="<%= formId %>ParentFormId"><%= parentFormId %></div>
 	<div class="hidden" id="<%= formId %>KeyFieldName"><%= keyFieldName %></div>
 	<div class="hidden" id="<%= formId %>ChildFormIds"><%= childFormIds %></div>
 	<div class="hidden" id="<%= formId %>DependantFormIds"><%= dependantFormIds %></div>
  <div class="hidden" id="<%= formId %>ParentFormIsDependant"><%= parentFormIsDependant %></div>
 	<div class="hidden" id="<%= formId %>AutoInsert"><%= autoInsert %></div>
 	<div class="hidden" id="<%= formId %>CascadeDeleteEnabled"><%= cascadeDeleteEnabled %></div>
  <div class="canvas formGridView <%= formGridViewCssClass %>" id="<%= formId %>GridView"></div>
  <div class="canvas formRecordView <%= formRecordViewCssClass %>" id="<%= formId %>RecordView"></div>
  <%
  if manualOrdering then
		%>
  	<div>
    <center>
      <input class="button" type="button" value="Subir" onclick="formMove('<%= formId %>', '<%= dbMoveRecordUp %>');">
      <input class="button" type="button" value="Bajar" onclick="formMove('<%= formId %>', '<%= dbMoveRecordDown %>');">
    </center>
  	</div>
    <%
  end if
end function

function doCallback(func)
  doCallback = true
  if len(func) = 0 then exit function
  dim ok: ok = eval(func)
  if failed then
    doCallback = reportError("Error interno en " & func & ".\n\nDetalles:\n\n" & errorMsg)
    if dbLogging then dbLog(func & " - FAILED: " & errorMsg)
  else
    doCallback = ok
    if dbLogging then
      if ok then 
        dbLog(func & " - OK")
      else
        dbLog(func & " - NOT OK:" & JSONResponse)
      end if
    end if
  end if
end function

function reportError(message)
  reportError = false
  JSONAddOpFailed
  JSONAddMessage(message)
  JSONSend
end function

dim inTransaction: inTransaction = false

function beginTransaction(ok)
  if inTransaction or not ok then exit function
  dbExecute("BEGIN TRANSACTION")
  inTransaction = true
end function

function finishTransaction(ok)
  if not inTransaction then exit function
  inTransaction = false
  if ok then
    dbExecute("COMMIT TRANSACTION")
  else
    dbExecute("ROLLBACK TRANSACTION")
  end if
end function

function logError
  dim sql
  sql = "INSERT INTO ERRORES (USUARIO, URL, METODO, RESPUESTA, SQL) VALUES (" & _
    join(array(sqlValue(getStringParam("usrName", 100)), sqlValue(getStringParam("url", 2000)), _
      sqlValue(getStringParam("method", 50)), sqlValue(getStringParam("response", 50000)), _
      sqlValue(replace(session("lastSQLCommand"), chr(39), chr(39) & chr(39)))), ",") & ")"
  dbGetData("SET NOCOUNT ON " & sql & " SELECT ID=SCOPE_IDENTITY() SET NOCOUNT OFF")
  response.write(rs("ID"))
  dbReleaseData
end function

function logFormActivity(params)
  if len(formTitle) = 0 then
    dim s
    s = replace(formId, "form", "")
    dim i, c
    dim t: t = ""
    for i = 1 to len(s)
      c = mid(s, i, 1)
      if i > 1 and c = uCase(c) then t = t & " " 
      t = t & c
    next
    logUserActivity t, params
  else
    logUserActivity formTitle, params
  end if
end function

function getFormRecordInfo
  dim i
  i = getFieldIndex("NOMBRE")
  if i < 0 then
    if recordId < 0 then
      getFormRecordInfo = "Inserting"
    else
      getFormRecordInfo = "ID=" & recordId
    end if
  elseif uBound(fieldNewValues) > 0 then
    getFormRecordInfo = dQuotes(fieldNewValues(i))
  elseif uBound(fieldCurrentValues) > 0 then
    getFormRecordInfo = dQuotes(fieldCurrentValues(i))
  else
    getFormRecordInfo = "ID=" & recordId
  end if
end function

dim b: b = dbConnect
if dbLogging then dbLog(formId & " - " & verb)
select case verb
  case "adminLoginDialog": renderAdminLoginDialog
  case "login": login
  case "getSessionInfo": getSessionInfo
  case "logout": logout
  case "sendPasswordRemainder": sendPasswordRemainder
  case "error": logError
  case else:
    if not handleCustomVerbs then
      response.write(getLoggedUsrData)
      if getLoggedUsrData then
        inserting = (verb = "newRecordView")
        eval(formId)
        if appReadOnly then usrAccessLevel = usrPermissionReadOnly
        if usrAccessLevel > usrPermissionNone then
          select case verb
            ' UI verbs
            case "formContainer":
              logFormActivity "open"
              renderFormsContainer
            case "form": eval(formRenderFunc)
            case "gridView": renderGridView
            case "recordView":
              if uBound(keyValues) >= 0 then
                recordViewDisabled = (cLng(keyValues(0)) = -1)
              end if
              renderRecordView
              if parentFormId = "none" and getNumericParam("selectedByUser") = 1 then logFormActivity "select " & getFormRecordInfo
            case "newRecordView":
              if recordViewIdFieldIsIdentity then
                recordViewFieldDefaults(getFieldIndex("ID")) = "auto"
              else
                select case suggestIDBy
                  case 10: recordViewFieldDefaults(getFieldIndex("ID")) = ((dbGetLastId(formTable, keyFieldNames, keyFieldValues) \ 10) + 1) * 10
                  case 1: recordViewFieldDefaults(getFieldIndex("ID")) = dbGetLastId(formTable, keyFieldNames, keyFieldValues) + 1
                end select
              end if
              renderRecordView
            case "fileUploadDialog": renderFileUploadDialog
            case "htmlEditorDialog": renderHTMLEditorDialog
            ' reporting verbs
            case "report": renderReport
            case "print": print
            ' data manipulation verbs
            case "binaryData": dbGetBinaryData formTable, keyFieldNames, keyFieldValues, recordId, false
            case "binaryFile": dbGetBinaryData formTable, keyFieldNames, keyFieldValues, recordId, true
            case "update", "delete", "deleteAll", "move":
              dim ok: ok = true
              select case verb
                case "update":
                  inserting = nullRecord
                  if inserting and usrAccessLevel < usrPermissionInsert then
                    reportError("No tiene permiso para agregar estos datos.")
                  elseif not inserting and usrAccessLevel < usrPermissionUpdate then
                    reportError("No tiene permiso para modificar estos datos.")
                  else
                    getNewValues
                    ok = doCallback(formBeforeUpdateFunc)
                    beginTransaction(ok)
                    if ok then
                      recordId = dbUpdate(formTable, keyFieldNames, keyFieldValues, recordViewFields, fieldChanged, fieldNewValues, _
                        recordViewReadOnlyFields, recordId, recordViewIdFieldIsIdentity, useAuditData)
                      ok = (recordId >= 0)
                    end if
                    nullRecord = (recordId < 0)
                    if ok then ok = doCallback(formAfterUpdateFunc)
                    finishTransaction(ok)
                    JSONSend
                  end if
                case "delete":
                  if usrAccessLevel < usrPermissionDelete then
                    reportError("No tiene permiso para eliminar estos datos.")
                  else
                    ok = doCallback(formBeforeDeleteFunc) 
                    beginTransaction(ok)
                    if ok then ok = dbDelete(formTable, keyFieldNames, keyFieldValues, recordId)
                    if ok then ok = doCallback(formAfterDeleteFunc)
                    finishTransaction(ok)
                    JSONSend
                  end if
                case "deleteAll":
                  if usrAccessLevel < usrPermissionDelete then
                    reportError("No tiene permiso para eliminar estos datos.")
                  else
                    eval(formPrepareQueryOptionsFunc)
                    dim searchExpr: searchExpr = getSearchExpr(dbSearchModeAll, keyFieldNames, keyFieldValues, null)
                    if searchExpr <> "" then searchExpr = " WHERE " & searchExpr
                    if dbGetData("SELECT ID FROM " & formTable & searchExpr) then
                      do while ok and not rs.EOF
                        ok = doCallback(formBeforeDeleteFunc)
                        beginTransaction(ok)
                        if ok then ok = dbDelete(formTable, keyFieldNames, keyFieldValues, rs("ID"))
                        if ok then ok = doCallback(formAfterDeleteFunc)
                        finishTransaction(ok)
                        rs.moveNext
                      loop
                      JSONSend
                    end if
                    dbReleaseData
                  end if
                case "move":
                  beginTransaction(ok)
                  ok = dbMoveRecord(getStringParam("direction", 5), formTable, keyFieldNames, keyFieldValues, recordId)
                  finishTransaction(ok)
              end select
              logFormActivity verb & " " & getFormRecordInfo
            case else:
              response.write("BAD_REQUEST (" & verb & ")")
          end select
        else
          response.write("NO_PERMISSION")
        end if
      else
        response.write("NO_SESSION")
      end if
    end if
end select
if b then dbDisconnect
%>


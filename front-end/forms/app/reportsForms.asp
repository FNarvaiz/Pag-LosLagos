
<!--#include file="reportRendering.asp"-->

<%
'    <link href="app/report.css" rel="stylesheet" type="text/css" media="screen">
'    <link href="app/reportPrint.css" rel="stylesheet" type="text/css" media="print">

dim reportTitle
dim reportFilename

function renderReport
  if not usrAccessAdminMaster and usrAccessLevel = usrPermissionNone then
    renderHTMLReport(reportAccessDenied)
  else
    dbGetData("SELECT * FROM INFORMES WHERE ID=" & recordId)
    reportTitle = rs("NOMBRE")
    dim func: func = rs("RENDERING_FUNCTION")
    dim format: format = rs("FORMATO")
    dim forDownload: forDownload = rs("PARA_DESCARGA")
    dim usrProfileRequired: usrProfileRequired = rs("ID_PERFIL_REQUERIDO")
    dbReleaseData
    if usrProfileRequired > usrProfile then
      renderHTMLReport(reportAccessDenied)
    else
      if isNull(func) then 
        func = "renderDefaultReport"
      else
        getNewValues
        dim ext: ext = ".html"
        select case format
          case "custom":
            eval(func)
          case "html": 
            renderHTMLReport(func)
          case "xls":
            renderExcelReport func, "application/vnd.ms-excel"
          case "text":
            renderPlainTextReport func, "text/plain"
          case else: 
            renderHTMLReport(reportFormatNotSupported)
        end select
        if forDownload then response.addHeader "Content-Disposition", "attachment; filename=" & reportFilename
      end if
    end if
  end if
end function

function print
  if not usrAccessAdminMaster and usrAccessLevel = usrPermissionNone then
    renderHTMLReport(reportAccessDenied)
  else
    dbConnect
    dbGetData("SELECT * FROM INFORMES WHERE ID=" & recordId)
    dim func: func = rs("RENDERING_FUNCTION")
    dim usrProfileRequired: usrProfileRequired = rs("ID_PERFIL_REQUERIDO")
    dbReleaseData
    if usrProfileRequired > usrProfile then
      renderHTMLReport(reportAccessDenied)
    else
      eval(func)
    end if
  end if
end function

function reportAccessDenied
  response.write("<p>No tiene permiso para ver este informe. Consulte con el administrador del sistema.</p>")
end function

function reportFormatNotSupported
  response.write("<p>Error: El formato " & format & " no está implementado.</p>")
end function

function renderDefaultReport
  response.write("El informe no áun no ha sido implementado.")
end function

function renderPlainTextReport(renderFunc, mimeType)
  response.contentType = mimeType
  eval(renderFunc)
end function

function renderExcelReport(renderFunc, mimeType)
  response.contentType = mimeType
  %>
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
  <%
  eval(renderFunc)
end function

function renderHTMLReport(renderFunc)
  %>
  <!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
  <html>
  <head>
  <title><%= reportTitle %> - Barrio de los Lagos - Sistema Brick</title>
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
  <style type="text/css" media="screen">
    <!--#include file="report.css"-->
  </style>
  <style type="text/css" media="print">
    <!--#include file="reportPrint.css"-->
  </style>
  <script type="text/javascript">
    <!--#include file="report.js"-->
  </script>
  </head>
  <body leftmargin="0" topmargin="0" >
    <div id="reportLogo"><img src="<%= formsAppResourcePath %>logo.png"></div>
    <div id="reportTitle"><%= reportTitle %></div>
    <div id="reportTimeStamp">Emisión: <%= zeroPad(day(date), 2) & "/" & zeroPad(month(date), 2) & "/" & year(date) %>&nbsp;
      <%= zeroPad(hour(time), 2) & ":" & zeroPad(minute(time), 2) %>&nbsp;por&nbsp;<%= usrName %></div>
    <%
      eval(renderFunc)
    %>
  </body>
  </html>
  <%
end function

function renderReportDateColumn(dateValue)
  if isNull(dateValue) then
    renderReportDateColumn = ""
  elseif year(dateValue) = thisYear then
    renderReportDateColumn = zeroPad(day(dateValue), 2) & dbDateSeparator & zeroPad(month(dateValue), 2)
  else
    renderReportDateColumn = zeroPad(day(dateValue), 2) & dbDateSeparator & zeroPad(month(dateValue), 2) & dbDateSeparator & _
      right(year(dateValue), 2)
  end if
end function

function renderReportFullDateColumn(dateValue)
  if isNull(dateValue) then
    renderReportFullDateColumn = ""
  else
    renderReportFullDateColumn = zeroPad(day(dateValue), 2) & dbDateSeparator & zeroPad(month(dateValue), 2) & dbDateSeparator & _
      year(dateValue)
  end if
end function

function renderReportHTMLValue(val)
  if isNull(val) then
    renderReportHTMLValue = "&nbsp;"
  elseif len(val) = 0 then
    renderReportHTMLValue = "&nbsp;"
  else
    renderReportHTMLValue = val
  end if
end function

function renderReportHTMLDecimal(decNum, decimals)
  if isNull(decNum) then
    renderReportHTMLDecimal = "&nbsp;"
  else
    renderReportHTMLDecimal = formatDecimal(round(decNum, decimals), decimals)
  end if
end function

function renderReportPDFDecimal(decNum, decimals)
  if isNull(decNum) then
    renderReportPDFDecimal = "-"
  else
    renderReportPDFDecimal = formatDecimal(round(decNum, decimals), decimals)
  end if
end function


' Form functions ==============================================================

function renderViewReportButton
  dim buttonLabel
  if fieldCurrentValues(getFieldIndex("PARA_DESCARGA")) then
    buttonLabel = "Descargar el informe"
  else
    buttonLabel = "Ver el informe"
  end if
  %>
  <input type="hidden" name="verb" value="report">
  <input type="hidden" name="formId" value="formReports">
  <input type="hidden" name="sessionId" value="<%= sessionId %>">
  <input type="hidden" name="recordId" value="<%= recordId %>">
  <input class="button" type="button" value="<%= buttonLabel %>"
    style="left: <%= recordViewLabelLeftPos %>px; top: <%= recordViewFieldTopPos %>px; min-width: 90px; padding: 0"
    onclick="normalPost('<%= formId %>', '<%= formsServer %>', '_blank')">
  <%
  renderViewReportButton = recordViewDefaultFieldHeight + 8
end function

function formReportRenderFunc
  usrAccessLevel = usrPermissionDelete
  renderDBFormControls
end function

function formAdminReports
  usrAccessLevel = usrAccessAdminReports
  formContainerCssClass = "reportsFormContainer"
  formContainerTitleCssClass = "reportsFormContainerTitle"
  formTitle = "Informes"
  forms = array("formReports")
  childFormIds = "formReports"
end function

function formReports
  usrAccessLevel = usrAccessAdminReports
  formTitle = ""
  formTable = "dbo.INFORMES_DE_USUARIO(" & usrId & ")"
  useAuditData = false

  formGridViewRowCount = 23
  formGridColumns = array("ID", "NOMBRE")
  formGridColumnTypes = array(formGridColumnHidden, formGridColumnGeneral)
  formGridColumnWidths = array(0, 350)
  gridViewOrderBy = "ID"
  gridViewReordering = false

  recordViewFieldLeftPos = 110
  recordViewEditboxWidth = 170
  recordViewLookupFieldNameField = "NOMBRE"

  if (verb = "recordView") or (verb = "report") then
    dim func: func = ""
    if dbGetData("SELECT RENDERING_FUNCTION FROM INFORMES WHERE ID=" & recordId) then
      func = rs("RENDERING_FUNCTION")
    end if
    dbReleaseData
    if len(func) = 0 then
      formRecordViewRenderFunc = null
    else
      formRecordViewRenderFunc = "formReportRenderFunc"
      select case func
        case "renderBookingListing":
          recordViewSeparators = array(null, "Condiciones del informe", null)
          recordViewFields = array("PARA_DESCARGA", "FECHA", "ID_RECURSO")
          recordViewDBFields = array(true, false, false)
          recordViewReadOnlyFields = array(true, false, false)
          recordViewFieldLabels = array("", "Fecha", "Recurso")
          recordViewFieldDefaults = array(null, date, null)
          recordViewNullableFields = array(true, false, true)
          recordViewFieldRenderFuncs = array("renderViewReportButton", "renderRecordViewDateField", _
            "renderRecordViewLookupField(" & dQuotes("RESERVAS_RECURSOS,ID") & ")")
        case else
          recordViewFields = array("PARA_DESCARGA")
          recordViewDBFields = array(true)
          recordViewReadOnlyFields = array(true)
          recordViewFieldLabels = array("")
          recordViewFieldDefaults = array(null)
          recordViewNullableFields = array(true)
          recordViewFieldRenderFuncs = array("renderViewReportButton")
      end select
    end if
  end if
end function

%>

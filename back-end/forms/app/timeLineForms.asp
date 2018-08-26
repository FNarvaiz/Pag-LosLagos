<%

function formAdminTimeLine
  usrAccessLevel = usrAccessAdminTimeLine
  formContainerCssClass = "timeLineFormContainer"
  formContainerTitleCssClass = "timeLineFormContainerTitle"
  formTitle = "Archivo Histórico"
  forms = array("formTimeLine")
  childFormIds = "formTimeLine"
end function

dim formTimeLineQueryTypeNames
formTimeLineQueryTypeNames = array("Todo", "Sin Aprobar")

dim formTimeLineQueryTypeSearchConditions
formTimeLineQueryTypeSearchConditions = array("", "APROBADO=0")

dim formTimeLineQuerySearchFieldNames
formTimeLineQuerySearchFieldNames = array("")

dim formTimeLineQuerySearchFields
formTimeLineQuerySearchFields = array("PERIODO + ' ' + NOMBRE")

dim formTimeLineQueryLimitNames
formTimeLineQueryLimitNames = array("hasta 100 ítems", "hasta 1000 ítems", "Todo")

dim formTimeLineQueryLimitClauses
formTimeLineQueryLimitClauses = array("TOP 100", "TOP 1000", "")

function renderFormTimeLineNotifyNeighborsBtn
  if nullRecord then exit function
  recordViewFieldTopPos = recordViewFieldTopPos + 4
  %>
  <input type="button" class="navButton" value="Enviar e-mail ahora"
    onclick="formTimeLineNotifyNeighbors(<%= recordId %>)"
    <%= renderBooleanAttr("disabled", recordViewDisabled or nullRecord) %>
    style="left: 10px; top: <%= recordViewFieldTopPos %>px; width: 120px"
  <%
  renderFormTimeLineNotifyNeighborsBtn = recordViewDefaultFieldHeight + 4
end function

function notifyNeighbors
  if dbGetData("SELECT NOMBRE, dbo.NOMBRE_CATEGORIA_LINEA_TIEMPO(ID_CATEGORIA) AS CATEGORIA, APROBADO, COALESCE(CONTENIDO_TEXTO, '') AS TEXTO, " & _
      "dbo.EMAILS_VECINOS() AS EMAILS, ARCHIVO_FILENAME, COALESCE(ARCHIVO_FILESIZE, 0) AS SIZE, ARCHIVO_CONTENTTYPE, ARCHIVO_BINARYDATA " & _
      "FROM LINEA_TIEMPO WHERE ID=" & recordId) then
    if not rs("APROBADO") then
      JSONAddMessage("Para enviar una notificación por e-mail, el ítem seleccionado debe estar aprobado.")
    elseif rs("SIZE") > 300 * 1024 then
      JSONAddMessage("El archivo es demasiado grande para enviarlo por mail a todos los vecinos. El máximo es 300 KB.")
    else
      dim invalidEmailAddresses: invalidEmailAddresses = validateEmailAddresses(rs("EMAILS"))
      dim result: result = sendCDOGroupMail(mainEmailAddress, rs("EMAILS").value, rs("NOMBRE").value, rs("TEXTO").value, _
        rs("ARCHIVO_FILENAME").value, rs("ARCHIVO_CONTENTTYPE").value, rs("ARCHIVO_BINARYDATA").value)
      if len(result) = 0 then
        if len(invalidEmailAddresses) = 0 then
          JSONAddMessage("Se ha enviado un e-mail a todos los vecinos, con copia a " & mainEmailAddress & ".")
        else
          JSONAddMessage("Se ha enviado un e-mail a todos los vecinos, con copia a " & mainEmailAddress & "." & _
            "\n\nSe han encontrado las siguientes direcciones de e-mail no válidas:\n\n" & replace(invalidEmailAddresses, ",", "\n"))
        end if
      else
        JSONAddMessage("Se ha producido un error al enviar el e-mail (Error: " & result & ").")
      end if
    end if
  else
    JSONAddMessage("Error interno: no se encontró el registro en el Archivo histórico (ID=" & recordId & ").")
  end if
  dbReleaseData
  JSONSend
end function

function renderFormTimeLineRecordView
  if usrProfile = usrProfileCommission then
    recordViewReadOnlyFields(getFieldIndex("APROBADO")) = true
  end if
  renderStandardRecordView
end function

function formTimeLineBeforeUpdate
  formTimeLineBeforeUpdate = true
  if not isNumeric(fieldNewValues(getFieldIndex("MES"))) then
    formTimeLineBeforeUpdate = reportError("El mes no es válido.")
    exit function
  end if
  if cInt(fieldNewValues(getFieldIndex("MES"))) < 1 or cInt(fieldNewValues(getFieldIndex("MES"))) > 12 then
    formTimeLineBeforeUpdate = reportError("El mes no es válido.")
    exit function
  end if
  if not isNumeric(fieldNewValues(getFieldIndex("ANIO"))) then
    formTimeLineBeforeUpdate = reportError("El año no es válido.")
    exit function
  end if
  if cInt(fieldNewValues(getFieldIndex("ANIO"))) < year(date) - 1 then
    formTimeLineBeforeUpdate = reportError("El Año no puede ser anterior a " & year(date) - 1 & ".")
    exit function
  end if
  if cInt(fieldNewValues(getFieldIndex("ANIO"))) > year(date) + 1 then
    formTimeLineBeforeUpdate = reportError("El Año no puede ser mayor a " & year(date) + 1 & ".")
    exit function
  end if
  'if not inserting then
   '' if len(fieldNewValues(getFieldIndex("ARCHIVO_FILENAME"))) > 0 and len(fieldNewValues(getFieldIndex("CONTENIDO_TEXTO"))) > 0 then
    ''  formTimeLineBeforeUpdate = reportError("ATENCION: El contenido no debe ser Texto y Archivo al mismo tiempo.\n\nSe mostrará el texto solamente.")
   ''   exit function
 ''   end if
  'end if
end function

function formTimeLine
  usrAccessLevel = usrAccessAdminTimeLine
  formTitle = ""
  formTable = "LINEA_TIEMPO"
  
  'formRecordViewRenderFunc = "renderFormTimeLineRecordView"
  formBeforeUpdateFunc = "formTimeLineBeforeUpdate"
  
  formGridViewRowCount = 26
  formGridColumns = array("PERIODO", "dbo.NOMBRE_CATEGORIA_LINEA_TIEMPO(ID_CATEGORIA)", _
    "NOMBRE", "CASE WHEN CONTENIDO_TEXTO IS NOT NULL THEN '(texto)' ELSE CAST(ARCHIVO_FILESIZE / 1024 AS VARCHAR) + '&nbsp;KB' END", _
    "CASE APROBADO WHEN 1 THEN 'hot' ELSE '' END")
  formGridColumnTypes = array(formGridColumnGeneralCenter, formGridColumnGeneralCenter, _
    formGridColumnGeneral, formGridColumnGeneralCenter, formGridColumnHidden)
  formGridColumnLabels = array("Período", "Categoría", "Descripción", "Tamaño", "")
  formGridColumnWidths = array(50, 120, 370, 60, 0)
  gridViewOrderBy = "PERIODO_FECHA DESC, ID_CATEGORIA, ID"
  formGridRowCssClassColumn = uBound(formGridColumns)
  gridViewReordering = false
  dim i
  dbGetData("SELECT COUNT(*) AS QTY FROM CATEGORIAS_LINEA_TIEMPO")
  i = rs("QTY")
  dbReleaseData
  if i > 0 then
    dim j: j = uBound(formTimeLineQueryTypeNames)
    redim preserve formTimeLineQueryTypeNames(i + j)
    redim preserve formTimeLineQueryTypeSearchConditions(i + j)
    dbGetData("SELECT * FROM CATEGORIAS_LINEA_TIEMPO ORDER BY ID")
    i = j + 1
    do while not rs.EOF
      formTimeLineQueryTypeNames(i) = "Categoría: " & rs("NOMBRE")
      formTimeLineQueryTypeSearchConditions(i) = "ID_CATEGORIA=" & rs("ID")
      rs.moveNext
      i = i + 1
    loop
    dbReleaseData
  end if

  recordViewFieldLeftPos = 80
  recordViewEditboxWidth = 240
  recordViewComboboxWidth = 120
  recordViewIdFieldIsIdentity = true
  
  recordViewSeparators = array(null, null, null, null, null, null, "Contenido", null, null, "Noficar a los vecinos")
  recordViewFields = array("ID", "ID_CATEGORIA", "MES", "ANIO", "NOMBRE", "APROBADO", "CONTENIDO_TEXTO", _
    "ARCHIVO_FILENAME", "ARCHIVO_FILESIZE", "NOTIFICAR")
  recordViewDBFields = array(true, true, true, true, true, true, true, true, true, false)
  recordViewReadOnlyFields = array(true, false, false, false, false, false, false, true, true, true)
  recordViewFieldLabels = array("Código", "Categoría", "Mes", "Año", "Descripción", "Aprobado", "Texto", "Archivo", "", "")
  recordViewFieldDefaults = array(null, null, month(date), year(date), null, false, null, null, null, null)
  recordViewNullableFields = array(false, false, false, false, false, true, true, true, true)
  recordViewFieldRenderFuncs = array("renderRecordViewIdentityField", _
    "renderRecordViewLookupField(" & dQuotes("CATEGORIAS_LINEA_TIEMPO,ID,VIGENTE=1 OR ID=dbo.ID_CATEGORIA_LINEA_TIEMPO(" & recordId & ")") & ")", _
    "renderRecordViewVarNumericField(2)", "renderRecordViewVarNumericField(4)", _
    "renderRecordViewLiteralField", "renderRecordViewBooleanField", "renderRecordViewTextField(5)", _
    "renderRecordViewFilenameField", "renderRecordViewHiddenField", _
    "renderFormTimeLineNotifyNeighborsBtn")

end function

%>

<%

function formAdminResourceBookings
  usrAccessLevel = usrAccessAdminBookings
  formContainerCssClass = "resourceBookingsFormContainer"
  formContainerTitleCssClass = "resourceBookingsFormContainerTitle"
  formTitle = "Reservas"
  forms = array("formResourceBookings")
  childFormIds = "formResourceBookings"
end function

dim formResourceBookingsQueryTypeNames
formResourceBookingsQueryTypeNames = array("Todas las reservas", "Reservas con resultado pendiente", "Reservas no utilizadas")

dim formResourceBookingsQueryTypeSearchConditions
formResourceBookingsQueryTypeSearchConditions = array("", "dbo.RESULTADO_RESERVA(ID) = 'Pendiente'", "ID_RESULTADO=20")

dim formResourceBookingsQuerySearchFieldNames
formResourceBookingsQuerySearchFieldNames = array("Por Nombre")

dim formResourceBookingsQuerySearchFields
formResourceBookingsQuerySearchFields = array("NOMBRE")

dim formResourceBookingsQueryLimitNames
formResourceBookingsQueryLimitNames = array("hasta 100 ítems", "hasta 1000 ítems", "Todo")

dim formResourceBookingsQueryLimitClauses
formResourceBookingsQueryLimitClauses = array("TOP 100", "TOP 1000", "")

function formResourceBookingsBeforeUpdate
  formResourceBookingsBeforeUpdate = true
  if not inserting then exit function

  dim neighborId: neighborId = fieldNewValues(getFieldIndex("ID_VECINO"))
  dim resourceId: resourceId = fieldNewValues(getFieldIndex("ID_RECURSO"))
  dim bookingType: bookingType = fieldNewValues(getFieldIndex("ID_TIPO"))
  dim bookingDate: bookingDate = stripDatePrefix(fieldNewValues(getFieldIndex("FECHA")))
  dim turnStart: turnStart = fieldNewValues(getFieldIndex("INICIO"))
  dim turnDuration: turnDuration = fieldNewValues(getFieldIndex("DURACION"))
  if not isNumeric(turnStart) then
    formResourceBookingsBeforeUpdate = reportError("El comienzo indicado no es válido.")
  end if
  turnStart = cInt(turnStart)
  if turnStart <= 0 or turnStart > 24 * 60 then
    formResourceBookingsBeforeUpdate = reportError("El comienzo indicado no es válido.")
    exit function
  end if
  if not isNumeric(turnDuration) then
    formResourceBookingsBeforeUpdate = reportError("La duración indicada no es válida.")
    exit function
  end if
  turnDuration = cInt(turnDuration)
  if turnDuration <= 0 or turnDuration > 24 * 60 then
    formResourceBookingsBeforeUpdate = reportError("La duración indicada no es válida.")
    exit function
  end if
  dbGetData("SELECT dbo.NOMBRE_RECURSO_RESERVA(" & resourceId & ")")
  dim resourceName: resourceName = rs(0)
  dbReleaseData
  if isNull(resourceName) then
    formResourceBookingsBeforeUpdate = reportError("El recurso indicado no es válido.")
    exit function
  end if
  dim ok
  dbGetData("SELECT dbo.TURNO_RESERVA_VALIDO(" & resourceId & ", " & turnStart & ", " & turnDuration & ")")
  ok = rs(0)
  dbReleaseData
  if not ok then
    formResourceBookingsBeforeUpdate = reportError(resourceName & ": el turno indicado no es válido.")
    exit function
  end if
  dbGetData("SELECT CAST(CASE WHEN dbo.FECHA_INICIO_TURNO(" & sqlValue(fieldNewValues(getFieldIndex("FECHA"))) & ", " & _
    turnStart + turnDuration & ") > GETDATE() THEN 1 ELSE 0 END AS BIT)")
  ok = rs(0)
  dbReleaseData
  if not ok then
    formResourceBookingsBeforeUpdate = reportError("No está permitido cargar turnos que ya han finalizado.")
    exit function
  end if
  dbGetData("SELECT dbo.TURNO_RESERVA_DISPONIBLE(" & resourceId & ", " & sqlValue(fieldNewValues(getFieldIndex("FECHA"))) & ", " & _
    turnStart & ", " & turnDuration & ")")
  ok = rs(0)
  dbReleaseData
  if not ok then
    if isNull(neighborId) then
      formResourceBookingsBeforeUpdate = reportError(resourceName & ": no se puede colocar un bloqueo para la fecha, comienzo y duración indicadas.")
    else
      formResourceBookingsBeforeUpdate = reportError(resourceName & ": no hay disponibilidad para el turno indicado.")
    end if
    exit function
  end if
  dbGetData("SELECT REQUIERE_VECINO_2 FROM RESERVAS_TIPOS WHERE ID_RECURSO = " & resourceId & " AND ID = " & bookingType)
  dim neighbor2Required: neighbor2Required = rs(0)
  dbReleaseData
  if neighbor2Required then
    if isNull(fieldNewValues(getFieldIndex("ID_VECINO_2"))) then
      formResourceBookingsBeforeUpdate = reportError(resourceName & ": el turno indicado requiere indicar el Vecino Extra.")
      exit function
    elseif fieldNewValues(getFieldIndex("ID_VECINO_2")) = neighborId then
      formResourceBookingsBeforeUpdate = reportError(resourceName & ": el Vecino Extra no puede ser titular del turno.")
      exit function
    end if
  else
    setFieldNewValue "ID_VECINO_2", null
  end if
end function

function renderFormResourceBookingsRecordView
  if isNull(fieldCurrentValues(getFieldIndex("ID_VECINO"))) then
    recordViewFieldRenderFuncs(getFieldIndex("ID_RESULTADO")) = "renderRecordViewHiddenField"
    recordViewFieldRenderFuncs(getFieldIndex("dbo.RESULTADO_RESERVA(ID)")) = "renderRecordViewLiteralField"
  else
    recordViewReadOnlyFields(getFieldIndex("ID_RESULTADO")) = isNull(fieldCurrentValues(getFieldIndex("dbo.RESULTADO_RESERVA(ID)"))) and _
      isNull(fieldCurrentValues(getFieldIndex("ID_RESULTADO")))
  end if
  renderStandardRecordView
end function

function jBookingTypeOptions(neighborId, resourceId)
  dim b
  if isNull(resourceId) then
    JSONAddStr "bookingTypeOptions", ""
  elseif isNull(neighborId) then
    b = dbConnect
    dbGetData("SELECT ID, NOMBRE FROM RESERVAS_TIPOS WHERE ID_RECURSO=" & sqlValue(resourceId) & " AND ID = 1 ORDER BY ID")
    JSONAddArray "bookingTypeOptions", rs.getRows, array("id", "name")
    dbReleaseData
    if b then dbDisconnect
  else
    b = dbConnect
    dbGetData("SELECT ID, NOMBRE FROM RESERVAS_TIPOS WHERE ID_RECURSO=" & sqlValue(resourceId) & " AND ID > 1 ORDER BY ID")
    JSONAddArray "bookingTypeOptions", rs.getRows, array("id", "name")
    dbReleaseData
    if b then dbDisconnect
  end if
  JSONSend
end function

function jBookingStartOptions(neighborId, resourceId, bookingTypeId)
  if isNull(resourceId) or (not isNull(neighborId) and isNull(bookingTypeId)) then
    JSONAddStr "bookingStartOptions", ""
  elseif isNull(neighborId) or not isNull(bookingTypeId) then
    dim b: b = dbConnect
    dbGetData("SELECT DISTINCT ID, NOMBRE FROM RESERVAS_TURNOS(" & resourceId & ") ORDER BY ID")
    JSONAddArray "bookingStartOptions", rs.getRows, array("id", "name")
    dbReleaseData
    if b then dbDisconnect
  end if
  JSONSend
end function

function jBookingDurationOptions(neighborId, resourceId, bookingTypeId, bookingStart)
  if isNull(resourceId) or isNull(bookingStart) or (not isNull(neighborId) and isNull(bookingTypeId)) then
    JSONAddStr "bookingDurationOptions", ""
  elseif isNull(neighborId) then
    dbGetData("SELECT ID, NOMBRE FROM dbo.RESERVAS_DURACIONES_BLOQUEOS(" & resourceId & ") ORDER BY ID")
    JSONAddArray "bookingDurationOptions", rs.getRows, array("id", "name")
    dbReleaseData
  else
    dbGetData("SELECT ID, NOMBRE FROM dbo.RESERVAS_DURACIONES_POSIBLES(" & resourceId & ", " & bookingTypeId & ", " & bookingStart & ") ORDER BY ID")
    JSONAddArray "bookingDurationOptions", rs.getRows, array("id", "name")
    dbReleaseData
  end if
  JSONSend
end function



function formResourceBookings
  usrAccessLevel = usrAccessAdminBookings
  formTitle = ""
  formTable = "RESERVAS"
  
  formRecordViewRenderFunc = "renderFormResourceBookingsRecordView"
  formBeforeUpdateFunc = "formResourceBookingsBeforeUpdate"

  formGridViewRowCount = 25
  formGridColumns = array("FECHA", "dbo.NOMBRE_RECURSO_RESERVA(ID_RECURSO)", "dbo.NOMBRE_TURNO(INICIO, DURACION)", _
    "dbo.NOMBRE_VECINO(ID_VECINO)", "dbo.NOMBRE_VECINO(ID_VECINO_2)", "dbo.RESULTADO_RESERVA(ID)", _
    "CASE WHEN ID_VECINO IS NULL THEN 'freezed' ELSE CASE WHEN dbo.RESULTADO_RESERVA(ID) = 'Pendiente' THEN 'hot' ELSE '' END END")
  formGridColumnLabels = array("Fecha", "Recurso", "Turno", "Vecino", "Vecino<br>extra", "Resultado", "")
  formGridColumnTypes = array(formGridColumnDate, formGridColumnGeneralCenter, formGridColumnGeneralCenter, _
    formGridColumnGeneralCenter, formGridColumnGeneralCenter, formGridColumnGeneralCenter, formGridColumnHidden)
  formGridRowCssClassColumn = uBound(formGridColumns)
  formGridColumnWidths = array(70, 110, 90, 130, 100, 110, 0)
  gridViewOrderBy = "1 DESC, 2, 3"
  gridViewReordering = false

  recordViewFieldLeftPos = 90
  recordViewEditboxWidth = 190
  recordViewIdFieldIsIdentity = true

  dim b: b = not nullRecord' and usrProfile <> usrProfileIT
  recordViewFields = array("ID", "FECHA", "ID_VECINO", "ID_RECURSO", "ID_TIPO", "INICIO", "DURACION", "ID_VECINO_2", "ID_RESULTADO", "dbo.RESULTADO_RESERVA(ID)")
  recordViewDBFields = array(true, true, true, true, true, true, true, true, true, true)
  recordViewReadOnlyFields = array(true, b, b, b, b, b, b, b, nullRecord, true)
  recordViewFieldLabels = array("", "Fecha", "Vecino", "Recurso", "Tipo", "Comienzo", "Duración", "Vecino extra", "Resultado", "Resultado")
  recordViewFieldDefaults = array(null, date, null, null, null, null, null, null, null, null)
  recordViewNullableFields = array(true, false, true, false, false, false, false, true, false, true)
  recordViewFieldRenderFuncs = array("renderRecordViewIdentityField", "renderRecordViewDateField", _
    "renderRecordViewLookupField(" & dQuotes("VECINOS,NOMBRE,ID=dbo.ID_VECINO_RESERVA(" & recordId & ") OR (HABILITADO=1 AND PERMISO_SERVICIOS=1),bookingNeighborFieldChanged,ID_RECURSO") & ")", _
    "renderRecordViewLookupField(" & dQuotes("RESERVAS_RECURSOS,ID,,bookingResourceFieldChanged,ID_TIPO") & ")", _
    "renderRecordViewLookupField(" & dQuotes("dbo.TIPO_RESERVA(" & recordId & "),ID,,bookingTypeFieldChanged,INICIO") & ")", _
    "renderRecordViewLookupField(" & dQuotes("dbo.RESERVAS_TURNOS(dbo.ID_RECURSO_RESERVA(" & recordId & ")),ID,,bookingStartFieldChanged,DURACION") & ")", _
    "renderRecordViewLookupField(" & dQuotes("dbo.RESERVAS_DURACIONES(" & recordId & "),ID") & ")", _
    "renderRecordViewLookupField(" & dQuotes("VECINOS,NOMBRE") & ")", _
    "renderRecordViewLookupField(" & dQuotes("RESERVAS_RESULTADOS,ID") & ")", "renderRecordViewHiddenField")

  dim i, j
  dbGetData("SELECT COUNT(*) AS QTY FROM RESERVAS_RECURSOS")
  i = rs("QTY")
  dbReleaseData
  j = uBound(formResourceBookingsQueryTypeNames)
  if i > 0 then
    redim preserve formResourceBookingsQueryTypeNames(i + j)
    redim preserve formResourceBookingsQueryTypeSearchConditions(i + j)
    dbGetData("SELECT * FROM RESERVAS_RECURSOS ORDER BY ID")
    i = j + 1
    do while not rs.EOF
      formResourceBookingsQueryTypeNames(i) = rs("NOMBRE")
      formResourceBookingsQueryTypeSearchConditions(i) = "ID_RECURSO=" & rs("ID")
      rs.moveNext
      i = i + 1
    loop
    dbReleaseData
  end if
end function

%>

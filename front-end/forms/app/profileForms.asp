<%

function formAdminProfile
  usrAccessLevel = usrPermissionUnrestricted
  formContainerCssClass = "profileFormContainer"
  formContainerTitleCssClass = "profileFormContainerTitle"
  formTitle = "Vecinos"
  forms = array("formProfile", "formProfileVisitors", "formProfileEvents", "formProfileEventAttendees", "formProfilePets")
  childFormIds = "formProfile"
end function

function formProfile
  usrAccessLevel = usrPermissionUnrestricted
  formTitle = "Mi Contraseña Telefónica<br><br>Te será solicitada al llamar a la Guardia."
  formTable = "VECINOS"
  childFormIds = "formProfileVisitors,formProfileEvents,formProfilePets"
  
  formGridColumns = array("ID")
  formGridColumnTypes = array(formGridColumnGeneral)
  formGridColumnWidths = array(100)
  defaultQueryLimit = -1
  addSearchExpr("ID=" & usrId)

  recordViewFieldLeftPos = 140
  recordViewEditboxWidth = 140
  recordViewButtons = array(false, false, true)
  useAuditData = false

  recordViewFields = array("ID", "CLAVE_TELEFONICA")
  recordViewDBFields = array(true, true)
  recordViewReadOnlyFields = array(true, false)
  recordViewFieldLabels = array("Código", "Contraseña (4 dígitos)")
  recordViewFieldDefaults = array(null, null)
  recordViewNullableFields = array(true, false)
  recordViewFieldRenderFuncs = array("renderRecordViewIdentityField", "renderRecordViewNumericField")
end function

function formProfileVisitors
  usrAccessLevel = usrPermissionUnrestricted

  formTitle = "Mis Visitantes Frecuentes"
  formTable = "VECINOS_VISITANTES_FRECUENTES"
  parentFormId = "formProfile"
  keyFieldName = "ID_VECINO"
  
  formGridViewRowCount = 8
  formGridColumns = array("ID", "NOMBRE", "TIPO")
  formGridColumnTypes = array(formGridColumnHidden, formGridColumnGeneral, formGridColumnGeneral)
  formGridColumnLabels = array("", "Nombre", "Tipo/Relación")
  formGridColumnWidths = array(0, 140, 80)
  gridViewReordering = false
  gridViewOrderBy = "2"
  defaultQueryLimit = -1

  useAuditData = false
  recordViewIdFieldIsIdentity = true
  recordViewFieldLeftPos = 90
  recordViewEditboxWidth = 160
  
  recordViewFields = array("ID", "NOMBRE", "TIPO", "DNI", "VEHICULO")
  recordViewDBFields = array(true, true, true, true, true)
  recordViewReadOnlyFields = array(true, false, false, false, false)
  recordViewFieldLabels = array("", "Nombre", "Tipo/Relación", "DNI", "Patente vehíc.")
  recordViewFieldDefaults = array(null, null, null, null, null)
  recordViewNullableFields = array(true, false, false, false, false)
  recordViewFieldRenderFuncs = array("renderRecordViewIdentityField", "renderRecordViewLiteralField", "renderRecordViewLiteralField", _
    "renderRecordViewDNIField", "renderRecordViewLiteralField")
end function

function formProfileEventsBeforeUpdate
  formProfileEventsBeforeUpdate = true
  if dateDiff("d", date, systemDateFromNewValue(getFieldIndex("FECHA"))) < 0 then
    formProfileEventsBeforeUpdate = reportError("La fecha no es válida.")
    exit function
  end if
end function

function renderFormProfileEventsRecordView
  if not nullRecord then
    recordViewReadOnly = dateDiff("d", fieldCurrentValues(getFieldIndex("FECHA")), date) > 0
  end if
  renderStandardRecordView
end function

function formProfileEvents
  usrAccessLevel = usrPermissionUnrestricted

  formTitle = "Mis Eventos"
  formTable = "VECINOS_EVENTOS"
  parentFormId = "formProfile"
  keyFieldName = "ID_VECINO"
  childFormIds = "formProfileEventAttendees"
  
  formBeforeUpdateFunc = "formProfileEventsBeforeUpdate"
  formRecordViewRenderFunc = "renderFormProfileEventsRecordView"
  
  formGridViewRowCount = 7
  formGridColumns = array("ID", "FECHA", "NOMBRE", "LUGAR")
  formGridColumnTypes = array(formGridColumnHidden, formGridColumnDate, formGridColumnGeneral, formGridColumnGeneral)
  formGridColumnLabels = array("", "Fecha", "Evento", "Lugar")
  formGridColumnWidths = array(0, 60, 170, 110)
  gridViewReordering = false
  gridViewOrderBy = "2 DESC"
  defaultQueryLimit = -1

  useAuditData = false
  recordViewIdFieldIsIdentity = true
  recordViewFieldLeftPos = 70
  recordViewEditboxWidth = 270
  
  recordViewFields = array("ID", "FECHA", "NOMBRE", "LUGAR")
  recordViewDBFields = array(true, true, true, true)
  recordViewReadOnlyFields = array(true, false, false, false)
  recordViewFieldLabels = array("", "Fecha", "Evento", "Lugar")
  recordViewFieldDefaults = array(null, null, null, null)
  recordViewNullableFields = array(true, false, false, false)
  recordViewFieldRenderFuncs = array("renderRecordViewIdentityField", "renderRecordViewDateField", "renderRecordViewLiteralField", _
    "renderRecordViewLiteralField")
end function

function renderFormProfileEventAttendeesRecordView
  if dbGetData("SELECT FECHA FROM VECINOS_EVENTOS WHERE ID=" & getKeyValue("ID_EVENTO")) then
    if datediff("d", rs(0), date) > 0 then
      recordViewReadOnly = true
      recordViewButtons = array(false, false, false)
    end if
  end if
  renderStandardRecordView
end function

function formProfileEventAttendees
  usrAccessLevel = usrPermissionUnrestricted

  formTitle = "Concurrentes"
  formTable = "VECINOS_EVENTOS_CONCURRENTES"
  parentFormId = "formProfileEvents"
  keyFieldName = "ID_EVENTO"
  
  formRecordViewRenderFunc = "renderFormProfileEventAttendeesRecordView"
  
  formGridViewRowCount = 10
  formGridColumns = array("ID", "NOMBRE")
  formGridColumnTypes = array(formGridColumnHidden, formGridColumnGeneral)
  formGridColumnLabels = array("", "Nombre")
  formGridColumnWidths = array(0, 180)
  gridViewReordering = false
  gridViewOrderBy = "2"
  defaultQueryLimit = -1

  useAuditData = false
  recordViewIdFieldIsIdentity = true
  recordViewFieldLeftPos = 60
  recordViewEditboxWidth = 140
  
  recordViewFields = array("ID", "NOMBRE")
  recordViewDBFields = array(true, true)
  recordViewReadOnlyFields = array(true, false)
  recordViewFieldLabels = array("", "Nombre")
  recordViewFieldDefaults = array(null, null)
  recordViewNullableFields = array(true, false)
  recordViewFieldRenderFuncs = array("renderRecordViewIdentityField", "renderRecordViewLiteralField")
end function

function renderFormProfilePetsImageUploadField
  renderRecordViewFileUploadField
  %>
  <table id="formProfilePetsImage" width="100%" height="100%" cellpadding="0" cellspacing="0">
    <tr>
      <td align="center">
      <%
      if nullRecord then
        response.write("(espacio para foto)")
      elseif isNull(fieldCurrentValues(recordViewCurrentField)) then
        response.write("(Mascota sin foto)")
      else
        %>
        <img class="anchor" onclick="window.open(this.src)"
          onload="imgResizeToFit(this, 160, 126); this.style.display = 'inline';"
          src="<%= formsServer %>?sessionId=<%= sessionId %>&verb=binaryData&formId=<%= formId %>&recordId=<%= recordId %>&dbFieldBaseName=FOTO&t=<%= timer() %>"
            title="click para ampliar la foto" style="display: none">
        <%
      end if
      %>
      </td>
    </tr>
  </table>
  <%
  renderFormProfilePetsImageUploadField = 0
end function

function formProfilePets
  usrAccessLevel = usrPermissionUnrestricted

  formTitle = "Mis Mascotas"
  formTable = "VECINOS_MASCOTAS"
  parentFormId = "formProfile"
  keyFieldName = "ID_VECINO"
  
  formGridViewRowCount = 5
  formGridColumns = array("ID", "NOMBRE")
  formGridColumnTypes = array(formGridColumnHidden, formGridColumnGeneral)
  formGridColumnLabels = array("", "Nombre")
  formGridColumnWidths = array(0, 170)
  gridViewReordering = false
  gridViewOrderBy = "2"
  defaultQueryLimit = -1

  useAuditData = false
  recordViewIdFieldIsIdentity = true
  recordViewFieldLeftPos = 60
  recordViewEditboxWidth = 180
  recordViewComboboxWidth = 100
  
  recordViewFields = array("ID", "NOMBRE", "ID_TIPO", "RAZA", "FOTO_CONTENTTYPE")
  recordViewDBFields = array(true, true, true, true, true)
  recordViewReadOnlyFields = array(true, false, false, false, true)
  recordViewFieldLabels = array("", "Nombre", "Tipo", "Raza", "Foto")
  recordViewFieldDefaults = array(null, null, null, null, null)
  recordViewNullableFields = array(true, false, false, false, false, true)
  recordViewFieldRenderFuncs = array("renderRecordViewIdentityField", "renderRecordViewLiteralField", _
    "renderRecordViewEnumField(" & dQuotes("Perro,Gato,Otro") & ")", _
    "renderRecordViewLiteralField", "renderFormProfilePetsImageUploadField")
end function

%>

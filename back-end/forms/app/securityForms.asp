<%

function formAdminSecurity
  if usrProfile = usrProfileSecurity or usrAccessAdminMaster then usrAccessLevel = usrPermissionUnrestricted
  formContainerCssClass = "securityFormContainer"
  formContainerTitleCssClass = "securityFormContainerTitle"
  formTitle = "Seguridad"
  forms = array("formSecurityNeighbors", "formSecurityPets", "formSecurityVisitors")
  childFormIds = "formSecurityNeighbors,formSecurityVisitors"
end function

dim formSecurityNeighborsQuerySearchFieldNames
formSecurityNeighborsQuerySearchFieldNames = array("Por Nombre")

dim formSecurityNeighborsQuerySearchFields
formSecurityNeighborsQuerySearchFields = array("UNIDAD + ' ' + NOMBRE")

function formSecurityNeighbors
  if usrProfile = usrProfileSecurity or usrAccessAdminMaster then usrAccessLevel = usrPermissionUnrestricted
  formTitle = ""
  formTable = "VECINOS"
  childFormIds = "formSecurityPets"
  
  formGridViewRowCount = 15
  formGridColumns = array("ID", "dbo.SPACE_PAD(UNIDAD, 15)", "NOMBRE", "CLAVE_TELEFONICA")
  formGridColumnTypes = array(formGridColumnHidden, formGridColumnGeneralCenter, formGridColumnGeneral, formGridColumnGeneralCenter)
  formGridColumnLabels = array("", "Unidad", "Vecino", "Clave telefónica")
  formGridColumnWidths = array(0, 60, 150, 60)
  gridViewOrderBy = "2"
  gridViewReordering = false
  defaultQueryLimit = -1
  addSearchExpr("HABILITADO=1")

  recordViewFields = array("ID")
  recordViewDBFields = array(true)
  recordViewReadOnlyFields = array(true)
  recordViewFieldLabels = array("")
  recordViewFieldDefaults = array(null)
  recordViewNullableFields = array(true)
  recordViewFieldRenderFuncs = array("renderRecordViewIdentityField")
end function

function renderFormSecurityPetsImageField
  %>
  <table id="formProfilePetsImage" width="100%" height="100%" cellpadding="0" cellspacing="0" style="border: 1px solid #999999;">
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
          onload="imgResizeToFit(this, 160, 98); this.style.display = 'inline';"
          src="<%= formsServer %>?sessionId=<%= sessionId %>&verb=binaryData&formId=<%= formId %>&recordId=<%= recordId %>&dbFieldBaseName=FOTO&t=<%= timer() %>"
            title="click para ampliar la foto" style="display: none">
        <%
      end if
      %>
      </td>
    </tr>
  </table>
  <%
  renderFormSecurityPetsImageField = 0
end function

function renderFormSecurityPetsRecordView
  renderDBFormControls
end function

function formSecurityPets
  usrAccessLevel = usrPermissionUnrestricted

  formTitle = "Mascotas"
  formTable = "VECINOS_MASCOTAS"
  parentFormId = "formSecurityNeighbors"
  keyFieldName = "ID_VECINO"
  
  formRecordViewRenderFunc = "renderFormSecurityPetsRecordView"
  formGridViewRowCount = 3
  formGridColumns = array("ID", "NOMBRE", "dbo.NOMBRE_TIPO_MASCOTA(ID_TIPO) + ' ' + RAZA")
  formGridColumnTypes = array(formGridColumnHidden, formGridColumnGeneral, formGridColumnGeneral)
  formGridColumnLabels = array("", "Nombre", "Tipo y raza")
  formGridColumnWidths = array(0, 150, 120)
  formGridViewShowFooter = false
  gridViewReordering = false
  gridViewOrderBy = "2"
  defaultQueryLimit = -1

  useAuditData = false
  recordViewIdFieldIsIdentity = true
  recordViewFieldLeftPos = 60
  recordViewEditboxWidth = 180
  recordViewComboboxWidth = 100
  
  recordViewFields = array("ID", "FOTO_CONTENTTYPE")
  recordViewDBFields = array(true, true)
  recordViewReadOnlyFields = array(true, true)
  recordViewFieldLabels = array("", "Foto")
  recordViewFieldDefaults = array(null, null)
  recordViewNullableFields = array(true, true)
  recordViewFieldRenderFuncs = array("renderRecordViewIdentityField", "renderFormSecurityPetsImageField")
end function

dim formSecurityVisitorsQuerySearchFieldNames
formSecurityVisitorsQuerySearchFieldNames = array("Por Nombre")

dim formSecurityVisitorsQuerySearchFields
formSecurityVisitorsQuerySearchFields = array("LOTE + ' ' + FAMILIA + ' ' + NOMBRE + ' ' + COALESCE(VEHICULO, '')")

function formSecurityVisitors
  if usrProfile = usrProfileSecurity or usrAccessAdminMaster then usrAccessLevel = usrPermissionUnrestricted

  formTitle = "Visitantes del día"
  formTable = "VISITANTES_DEL_DIA"
  
  formGridViewRowCount = 27
  formGridColumns = array("ID", "dbo.SPACE_PAD(LOTE, 15)", "FAMILIA", "NOMBRE", "DNI", "VEHICULO", "EVENTO", "LUGAR")
  formGridColumnTypes = array(formGridColumnHidden, formGridColumnGeneralCenter, formGridColumnGeneral, formGridColumnGeneral, _
    formGridColumnGeneral, formGridColumnGeneral, formGridColumnGeneral, formGridColumnGeneral)
  formGridColumnLabels = array("", "Unidad", "Vecino", "Visitante", "DNI", "Patente", "Evento", "Lugar")
  formGridColumnWidths = array(0, 50, 100, 130, 60, 60, 120, 110)
  gridViewReordering = false
  gridViewOrderBy = "2"
  defaultQueryLimit = -1
  formGridViewSelectable = false
  addSearchExpr("DATEDIFF(DAY, GETDATE(), FECHA) = 0")

  useAuditData = false
  recordViewIdFieldIsIdentity = true
  recordViewFieldLeftPos = 60
  recordViewEditboxWidth = 180
  
  recordViewFields = array("ID", "NOMBRE", "DNI", "VEHICULO")
  recordViewDBFields = array(true, true, true, true)
  recordViewReadOnlyFields = array(true, false, false, false)
  recordViewFieldLabels = array("", "Nombre", "DNI", "Vehículo")
  recordViewFieldDefaults = array(null, null, null, null)
  recordViewNullableFields = array(true, false, false, false)
  recordViewFieldRenderFuncs = array("renderRecordViewIdentityField", "renderRecordViewLiteralField", "renderRecordViewDNIField", _
    "renderRecordViewLiteralField")
end function


%>

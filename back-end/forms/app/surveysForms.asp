<%

function formAdminSurveys
  usrAccessLevel = usrAccessAdminSurveys
  formContainerCssClass = "SurveysFormContainer"
  formContainerTitleCssClass = "SurveysFormContainerTitle"
  formTitle = "Encuestas"
  forms = array("formSurveys", "formSurveysReviews")
  childFormIds = "formSurveys,formSurveysReviews"
end function

dim formSurveysQueryTypeNames
formSurveysQueryTypeNames = array("Vigentes", "No vigentes", "Todos")

dim formSurveysQueryTypeSearchConditions
formSurveysQueryTypeSearchConditions = array("VIGENTE=1", "VIGENTE=0", "")

dim formSurveysQuerySearchFieldNames
formSurveysQuerySearchFieldNames = array("Por Nombre")

dim formSurveysQuerySearchFields
formSurveysQuerySearchFields = array("NOMBRE")

function formSurveysBeforeUpdate
  formSurveysBeforeUpdate =  true
  if fieldChanged(getFieldIndex("VIGENTE")) then
    if cInt(fieldNewValues(getFieldIndex("VIGENTE"))) = 0 then
      setFieldNewValue "FECHA_BAJA", date
    else
      setFieldNewValue "FECHA_BAJA", null
    end if
  end if
end function

function formSurveys
  usrAccessLevel = usrAccessAdminSurveys
  formTitle = ""
  formTable = "ENCUESTAS"
  childFormIds = "formSurveysReviews"
  
  formBeforeUpdateFunc = "formSurveysBeforeUpdate"
  
  formGridViewRowCount = 27
  formGridColumns = array("FECHA_ALTA", "NOMBRE", "dbo.NOMBRE_VALORACION(dbo.VALORACION_ENCUESTA(ID))", "dbo.CANTIDAD_VALORACIONES_ENCUESTA(ID)")
  formGridColumnTypes = array(formGridColumnDate, formGridColumnGeneral, formGridColumnGeneralCenter, formGridColumnGeneralCenter)
  formGridColumnLabels = array("Alta", "Titulo", "Valor.", "Votos")
  formGridColumnWidths = array(60, 220, 40, 40)
  gridViewOrderBy = "1 DESC"
  defaultQueryLimit = -1

  recordViewFieldLeftPos = 80
  recordViewEditboxWidth = 200
  recordViewIdFieldIsIdentity = true

  recordViewSeparators = array(null, null, null, null, null, null, "Descripción")
  recordViewFields = array("ID", "ID_CATEGORIA", "NOMBRE", "FECHA_ALTA", "VIGENTE", "FECHA_BAJA", "DESCRIPCION")
  recordViewDBFields = array(true, true, true, true, true, true, true, true)
  recordViewReadOnlyFields = array(true, false, false, true, false, true, false, true)
  recordViewFieldLabels = array("Código", "Categoría", "Nombre", "Fecha alta", "Vigente", "Fecha baja", "Descripción")
  recordViewFieldDefaults = array(null, null, null, null, true, null, null)
  recordViewNullableFields = array(true, false, false, true, false, true, false)
  recordViewFieldRenderFuncs = array("renderRecordViewIdentityField", _
    "renderRecordViewLookupField(" & dQuotes("CATEGORIAS_ENCUESTAS,ID,VIGENTE=1 OR ID=dbo.ID_CATEGORIA_ENCUESTA(" & recordId & ")") & ")", _
    "renderRecordViewLiteralField", "renderRecordViewDateField", "renderRecordViewBooleanField", "renderRecordViewDateField", _
    "renderRecordViewNotesField(10)")

  dim b: b = dbConnect
  dim i
  dbGetData("SELECT COUNT(*) AS QTY FROM CATEGORIAS_ENCUESTAS")
  i = rs("QTY")
  dbReleaseData
  if i > 0 then
    redim preserve formSurveysQueryTypeNames(i + 2)
    redim preserve formSurveysQueryTypeSearchConditions(i + 2)
    dbGetData("SELECT * FROM CATEGORIAS_ENCUESTAS ORDER BY ID")
    i = 3
    do while not rs.EOF
      formSurveysQueryTypeNames(i) = "Categoría: " & rs("NOMBRE")
      formSurveysQueryTypeSearchConditions(i) = "VIGENTE=1 AND ID_CATEGORIA=" & rs("ID")
      rs.moveNext
      i = i + 1
    loop
    dbReleaseData
  end if
  if b then dbDisconnect
end function

function renderSurveysReviewsRecordView
  renderDBFormControls
end function

function renderRecordViewLiteral2Field(maxLength)
  renderRecordViewFieldLabel
  renderCurrentValueField
  dim elemName: elemName = renderHTMLFieldNewValueName(recordViewCurrentField)
  %>
	<textarea class="textarea" name="<%= elemName %>" type="text" maxlength="<%= maxLength %>"
    style="left: <%= recordViewFieldLeftPos %>px; top: <%= recordViewFieldTopPos %>px; width: <%= recordViewEditboxWidth %>px;
      height: <%= 5 * recordViewDefaultFieldHeight - 2 %>px; <%= renderFieldBackgroundColor %>" 
      <%= renderBooleanAttr("disabled", recordViewDisabled) %>
		<%= renderValidationAttribute %> <%= renderBooleanAttr("readonly", isFieldReadOnly(recordViewCurrentField)) %>
    onFocus="this.select()"><%= renderHTMLFieldValue %></textarea>
  <%
  renderRecordViewLiteral2Field = 5 * recordViewDefaultFieldHeight
end function

function formSurveysReviews
  usrAccessLevel = usrAccessAdminSurveys
  formTitle = "Valoraciones de los vecinos"
  formTable = "ENCUESTAS_VALORACIONES"
  parentFormId = "formSurveys"
  keyFieldName = "ID_ENCUESTA"
  useAuditData = false
  
  formRecordViewRenderFunc = "renderSurveysReviewsRecordView" 
  
  formGridViewRowCount = 21
  formGridColumns = array("dbo.UNIDAD_VECINO(ID_VECINO)", "dbo.NOMBRE_VECINO(ID_VECINO)", "dbo.NOMBRE_VALORACION(VALORACION)")
  formGridColumnTypes = array(formGridColumnGeneralCenter, formGridColumnGeneral, formGridColumnGeneralCenter)
  formGridColumnLabels = array("Unidad", "Familia", "Valoración")
  formGridColumnWidths = array(60, 120, 70)
  gridViewOrderBy = "1"
  defaultQueryLimit = -1

  recordViewFieldLeftPos = 70
  recordViewEditboxWidth = 200

  recordViewFields = array("ID", "COMENTARIO")
  recordViewDBFields = array(true, true)
  recordViewReadOnlyFields = array(true, true)
  recordViewFieldLabels = array("Código", "Comentario")
  recordViewFieldDefaults = array(null, null)
  recordViewNullableFields = array(true, true)
  recordViewFieldRenderFuncs = array("renderRecordViewHiddenField", "renderRecordViewLiteral2Field(140)")
end function

%>

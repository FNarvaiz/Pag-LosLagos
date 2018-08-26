<%
' Form parametrization

function formAdminContents
  usrAccessLevel = usrAccessAdminContents

  formContainerCssClass = "contentsFormContainer"
  formContainerTitleCssClass = "contentsFormContainerTitle"
  formTitle = "Contenidos"
  forms = array("formContents")
  childFormIds = "formContents"
end function

function formContents
  usrAccessLevel = usrAccessAdminContents

  formTitle = ""
  formTable = "CONTENIDOS"
  
  formGridViewRowCount = 22
  formGridColumns = array("ID", "NOMBRE")
  formGridColumnTypes = array(formGridColumnGeneralCenter, formGridColumnGeneral)
  formGridColumnWidths = array(60, 280)
  formGridColumnLabels = array("Código", "Título")
  gridViewOrderBy = "1"

  recordViewFieldLeftPos = 60
  recordViewEditboxWidth = 200

  recordViewSeparators = array(null, null, null)
  recordViewFields = array("ID", "NOMBRE", "HTML")
  recordViewDBFields = array(true, true, true)
  recordViewReadOnlyFields = array(false, false, false)
  recordViewFieldLabels = array("Código", "Título", "Contenido")
  recordViewFieldDefaults = array(null, null, null)
  recordViewNullableFields = array(false, false, true)
  recordViewFieldRenderFuncs = array("renderRecordViewNumericField", "renderRecordViewLiteralField", "renderRecordViewHTMLField")
end function

%>

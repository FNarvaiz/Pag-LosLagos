<%

dim formsServerPath
formsServerPath = "https://" & Request.ServerVariables("SERVER_NAME") & left(request.serverVariables("URL"), inStrRev(request.serverVariables("URL"), "/"))
dim formsServer: formsServer = formsServerPath & "forms.asp"
dim formsAppPath: formsAppPath = formsServerPath & "app/"
dim formsAppResourcePath: formsAppResourcePath = formsAppPath & "resource/"
dim rootPath
if inStrRev(formsServerPath, "forms/") > 0 then
  rootPath = left(formsServerPath, inStrRev(formsServerPath, "forms/") - 1)
else
  rootPath = formsServerPath
end if

dim formInService: formInService = true
dim appReadOnly: appReadOnly = false

setLocale(11274)

const uiDecimalSeparator = ","
dim sysDecimalSeparator
if isNumeric("1,1") then
  sysDecimalSeparator = ","
else
  sysDecimalSeparator = "."
end if
const notAvailableFieldValue = "N/D"
const notApplicableFieldValue = "N/A"

' Input parameters from remote client
dim verb:             verb = getStringParam("verb", 50)
dim formId:           formId = getStringParam("formId", 50)
dim keyFieldNames:    keyFieldNames = getStringParam("keyFields", 100)
dim keyFields:        keyFields = strToArray(keyFieldNames, ",")
dim keyFieldValues:   keyFieldValues = getStringParam("keyValues", 200)
dim keyValues:        keyValues = strToArray(keyFieldValues, ",")
dim recordId:         recordId = getIntegerParam("recordId"): if isNull(recordId) then recordId = -1
dim queryOrder:       queryOrder = getStringParam("queryOrder", 20)
dim queryType:        queryType = getIntegerParam("queryType"): if isNull(queryType) then queryType = 0
dim querySearchField: querySearchField = getIntegerParam("querySearchField"): if isNull(querySearchField) then querySearchField = 0
dim querySearchValue: querySearchValue = getStringParam("querySearchValue", 50)
dim queryLimit:       queryLimit = getIntegerParam("queryLimit"): if isNull(queryLimit) then queryLimit = 0

' form structure definition data
dim forms:                 forms = null
dim formTitle:             formTitle = "Título"
dim formTable:             formTable = "TABLE"
dim childFormIds:          childFormIds = "none"
dim parentFormId:          parentFormId = "none"
dim dependantFormIds:      dependantFormIds = "none"
dim keyFieldName:          keyFieldName = "none"
dim searchCondition:       searchCondition = ""
dim suggestIDBy:           suggestIDBy = 10
dim useAuditData:          useAuditData = true
dim autoInsert:            autoInsert = false
dim literalKeyField:       literalKeyField = ""
dim parentFormIsDependant: parentFormIsDependant = false
dim cascadeDeleteEnabled:  cascadeDeleteEnabled = true


' Rendering callbacks
dim formRenderFunc:              formRenderFunc = "renderDBForm"
dim formGridViewRenderFunc:      formGridViewRenderFunc = "renderStandardGridView"
dim formRecordViewRenderFunc:    formRecordViewRenderFunc = "renderStandardRecordView"
dim formQueryOptionsRenderFunc:  formQueryOptionsRenderFunc = "renderStandardQueryOptions"

dim formPrepareQueryOptionsFunc: formPrepareQueryOptionsFunc = "standardPrepareQueryOptions"

' DB update callbacks
dim formBeforeDeleteFunc:     formBeforeDeleteFunc = ""
dim formBeforeUpdateFunc:     formBeforeUpdateFunc = ""
dim formAfterDeleteFunc:      formAfterDeleteFunc = ""
dim formAfterUpdateFunc:      formAfterUpdateFunc = ""

' Global css classes
dim formContainerCssClass:       formContainerCssClass = ""
dim formContainerTitleCssClass:  formContainerTitleCssClass = ""
dim formCssClass:                formCssClass = ""
dim formTitleCssClass:           formTitleCssClass = ""
dim formGridViewCssClass:        formGridViewCssClass = ""
dim formGridViewListboxCssClass: formGridViewListboxCssClass = ""
dim formRecordViewCssClass:      formRecordViewCssClass = ""

' Grid view data
const formGridColumnHidden = 0
const formGridColumnGeneral = 10
const formGridColumnGeneralCenter = 11
const formGridColumnName = 15
const formGridColumnNumber = 20
const formGridColumnBoolean = 30
const formGridColumnCurrency = 50
const formGridColumnCurrency4 = 51
const formGridColumnDecimal2 = 60
const formGridColumnDecimal3 = 61
const formGridColumnPercent = 70
const formGridColumnDate = 80
const formGridColumnTime = 90
const formGridColumnDateTime = 100
const formGridColumnImage = 200
const formGridColumnFileIcon = 220

const formGridRowHeight = 16

dim gridViewOrderBy:           gridViewOrderBy = "1"
dim gridViewReordering:        gridViewReordering = true
dim gridViewQueryLimitClause:  gridViewQueryLimitClause = ""
dim formGridViewRowCount:      formGridViewRowCount = 1
dim formGridColumns:           formGridColumns = null
dim formGridColumnTypes:       formGridColumnTypes = null
dim formGridColumnWidths:      formGridColumnWidths = null
dim formGridColumnLabels:      formGridColumnLabels = null
dim formGridTotals:            formGridTotals = null
dim formGridRowCssClassColumn: formGridRowCssClassColumn = null
dim manualOrdering:            manualOrdering = false
dim formGridViewSelectable:    formGridViewSelectable = true
dim formGridViewShowTools:     formGridViewShowTools = false
dim formGridViewShowFooter:    formGridViewShowFooter = true
dim gridViewDetailModule:      gridViewDetailModule = ""
dim defaultQueryLimit:         defaultQueryLimit = 100

' Record view data
dim recordViewButtons:               recordViewButtons = array(true, true, true)
dim recordViewLabelLeftPos:          recordViewLabelLeftPos = 8
dim recordViewFieldLeftPos:          recordViewFieldLeftPos = 100
dim recordViewTopmostFieldTopPos:    recordViewTopmostFieldTopPos = 8
dim recordViewEditboxWidth:          recordViewEditboxWidth = 100
dim recordViewDefaultFieldHeight:    recordViewDefaultFieldHeight = 20
dim recordViewNumericEditboxWidth:   recordViewNumericEditboxWidth = 50
dim recordViewComboboxWidth:         recordViewComboboxWidth = 0
dim recordViewDateEditboxWidth:      recordViewDateEditboxWidth = 70
dim recordViewTimeEditboxWidth:      recordViewTimeEditboxWidth = 55
dim recordViewDateTimeEditboxWidth:  recordViewDateTimeEditboxWidth = 101
dim recordViewTextAreaHeight:        recordViewTextAreaHeight = 58
dim fileUploadFormDataType:          fileUploadFormDataType = "imagen"
dim fileUploadFormMaxFilesize:       fileUploadFormMaxFilesize = 1024 * 1024

' Record view fields definitions
dim recordViewSeparators:            recordViewSeparators = null
dim recordViewFields:                recordViewFields = null
dim recordViewDBFields:              recordViewDBFields = null
dim recordViewReadOnlyFields:        recordViewReadOnlyFields = null
dim recordViewFieldLabels:           recordViewFieldLabels = null
dim recordViewFieldDefaults:         recordViewFieldDefaults = null
dim recordViewNullableFields:        recordViewNullableFields = null
dim recordViewFieldRenderFuncs:      recordViewFieldRenderFuncs = null
dim recordViewIdFieldIsIdentity:     recordViewIdFieldIsIdentity = false
dim recordViewLookupFieldNameField:  recordViewLookupFieldNameField = "NOMBRE"

' Permission levels
const usrPermissionNone = 0
const usrPermissionReadOnly = 1
const usrPermissionInsert = 2
const usrPermissionUpdate = 3
const usrPermissionDelete = 4
const usrPermissionUnrestricted = 5

' Record view render process variables
dim recordViewCurrentField:  recordViewCurrentField = 0
dim recordViewFieldTopPos:   recordViewFieldTopPos= 0
dim fieldCurrentValues:      fieldCurrentValues = array("")
dim fieldNewValues:          fieldNewValues = array("")
dim fieldChanged:            fieldChanged = array("")
dim recordViewDisabled:      recordViewDisabled = false
dim recordViewReadOnly:      recordViewReadOnly = false
dim usrAccessLevel:          usrAccessLevel = usrPermissionNone
dim nullRecord:              nullRecord = (recordId < 0)

' Audit info
dim auditUserName: auditUserName = ""
dim auditDate: auditDate = ""

' Data Update
dim inserting: inserting = false

%>


var server = "forms/forms.asp";
var gridRowIdSeparator = "@";
var firstFormRecordId = -1;

var currentFormContainerId = "";

function showDialog()
{
  document.getElementById("dialogbackgnd").style.visibility = "visible";
  document.getElementById("dialogsPanel").style.visibility = "visible";
}

function hideDialog()
{
  document.getElementById(loadingSignalId).style.visibility = "hidden";
  document.getElementById("dialogbackgnd").style.visibility = "hidden";
  document.getElementById("dialogsPanel").style.visibility = "hidden";
}

function isFormLoaded(formId)
{
  if (document.getElementById(formId))
	  return(true);
	else
	  return(false);
}

function isFormContainer(formId)
{
  var elem = document.getElementById(formId);
	return(elem && elem.className.indexOf("formContainer") >= 0);
}

function getParentFormId(formId)
{
  return(document.getElementById(formId + "ParentFormId").innerHTML);
}

function parentFormIsDependant(formId)
{
  var elem = document.getElementById(formId + "ParentFormIsDependant");
	return(elem && elem.innerHTML.toLowerCase() == "true");
}

function getChildFormIds(formId)
{
  return(document.getElementById(formId + "ChildFormIds").innerHTML.split(","));
}

function getDependantFormIds(formId)
{
  return(document.getElementById(formId + "DependantFormIds").innerHTML.split(","));
}

function getKeyFieldName(formId)
{
  return(document.getElementById(formId + "KeyFieldName").innerHTML);
}

function getFormDetailModule(formId)
{
  var elem = document.getElementById(formId + "DetailModule")
  if (elem)
    return(elem.innerHTML);
  else
    return(null);
}

function getRecordId(formId)
{
  var elem = document.getElementById(formId + "SelectedItem");
  if (elem)
    return(parseInt(elem.innerHTML));
  else
    return(-1);
}

function formsEmpty(formIds)
{
  for (var i = 0; i < formIds.length; i++)
    if (document.getElementById(formIds[i] + "IsEmpty").innerHTML.toLowerCase() != "true") return(false);
  return(true);
}

function cascadeDeleteEnabled(formId)
{
  return(document.getElementById(formId + "CascadeDeleteEnabled").innerHTML.toLowerCase() == "true");
}

function formAutoInsert(formId)
{
  return(document.getElementById(formId + "AutoInsert").innerHTML.toLowerCase() == "true");
}

function formCurrentRecordIsVisible(formId)
{
  return(document.getElementById(formId + "CurrentRecordIsVisible").innerHTML.toLowerCase() == "true");
}

function getQueryOrder(formId)
{
  var elem = document.getElementById(formId + "QueryOrder");
	if (elem)
    return(elem.innerHTML);
	else
    return("");
}

function setQueryOrder(formId, newQueryOrder)
{
  var elem = document.getElementById(formId + "QueryOrder");
	if (elem)
    elem.innerHTML = newQueryOrder;
}

function setRecordId(formId, recordId)
{
  document.getElementById(formId + "SelectedItem").innerHTML = recordId;
  if (recordId > 0)
    scrollSelectedItemIntoView(formId, true);
  inserting = false;
  insertingFormId = "";
}

function setLiteralKeyValue(formId)
{
  if (formIdList[0] == formId)
  {
    var labelElem = document.getElementById(currentFormContainerId + "SelectedKey");
    if (inserting)
    {
      if (insertingFormId == formId)
        labelElem.innerHTML = "&nbsp;&nbsp;-&nbsp;&nbsp;(agregando...)";
    }
    else
    {
      var frm = document.forms[formId];
      var literalKeyFieldHolderElem = document.getElementById(formId + "LiteralKeyField");
      if (literalKeyFieldHolderElem)
      {
        var literalKeyField = frm.elements["field" + literalKeyFieldHolderElem.innerHTML + "Value"];
        if (literalKeyField)
          labelElem.innerHTML = "&nbsp;&nbsp;-&nbsp;&nbsp;" + literalKeyField.value;
        else
          labelElem.innerHTML = "";
      }
      else
        labelElem.innerHTML = "";
    }
  }
}

function deselectGridViewItem(formId)
{
  var currentRecordId = getRecordId(formId);
	if (currentRecordId >= 0)
	{
    var elem = document.getElementById(formId + gridRowIdSeparator + currentRecordId);
		elem.className = "";
	}
}

function querySearchChanged(evt, formId)
{
  evt = (evt) ? evt : ((window.event) ? event : null);
  if (!evt) return;
  if (loadingForms) 
    cancelEvent(evt);
  if (evt.keyCode == 13)
    refreshForm(formId);
}

function querySearchGo(formId)
{
  if (!loadingForms) 
    refreshForm(formId);
}
function querySearchClear(formId)
{
  if (loadingForms) return;
  var e = document.getElementById(formId + "QuerySearchValue")
  if (e && e.value)
  {
    e.value = "";
    refreshForm(formId);
  }
}

function gridColumnClick(formId, colNumber)
{
  if (loadingForms) return;
  var currentQueryOrder = getQueryOrder(formId);
  if (currentQueryOrder == colNumber.toString())
    setQueryOrder(formId, colNumber + " DESC");
  else
    setQueryOrder(formId, colNumber);
  refreshForm(formId);
}

function getQueryParams(formId)
{
  var s = "";
  if (document.getElementById(formId + "QueryOptions"))
  {
    var searchFieldElem = document.getElementById(formId + "QuerySearchField");
    var searchValueElem = document.getElementById(formId + "QuerySearchValue");
    if (searchFieldElem && searchFieldElem.value && searchValueElem && searchValueElem.value)
      s = "&querySearchField=" + searchFieldElem.value + "&querySearchValue=" + encodeURIComponent(searchValueElem.value);
    var queryTypeElem = document.getElementById(formId + "QueryTypeCombobox");
    if (queryTypeElem && queryTypeElem.value)
      s += "&queryType=" + queryTypeElem.value;
    var queryLimitElem = document.getElementById(formId + "QueryLimitCombobox");
    if (queryLimitElem && queryLimitElem.value)
      s += "&queryLimit=" + queryLimitElem.value;
  }
  s += "&queryOrder=" + getQueryOrder(formId);
  return(s);
}

function buildKeyParams(formId)
{
  var keyFields = "";
  var keyValues = "";
  var mId = formId;
  var pId = getParentFormId(mId);
  while (pId != "none")
  {
    if (keyFields)
    {
      keyFields += ","
      keyValues += ","
    }
    keyFields += getKeyFieldName(mId);
    keyValues += getRecordId(pId);
    mId = pId;
    pId = getParentFormId(mId);
  }
  if (keyFields)
    return("&keyFields=" + keyFields + "&keyValues=" + keyValues);
  else
    return("")
}

function buildRequestUrl(verb, formId)
{
	var url = server + "?sessionId=" + sessionId + "&verb=" + verb + "&formId=" + formId;
  if (verb != "form")
  {
    if (verb == "gridView")
    {
      if(firstFormRecordId >= 0 && formId == formIdList[0])
        url += "&recordId=" + firstFormRecordId;
      else
        url += "&recordId=" + getRecordId(formId);
      url += getQueryParams(formId);
    }
    else
      url += "&recordId=" + getRecordId(formId);
    url += buildKeyParams(formId);
  }
  return(url);
}

function loadDetailModule(detailFormId, selectedRecordId)
{
  firstFormRecordId = selectedRecordId;
  loadFormContainer(detailFormId, document.getElementById(currentFormContainerId).parentNode.id);
}

function grc(trElem) // formGridRecordChanged
{
  if (loadingForms) return;
  var s = trElem.id.split(gridRowIdSeparator);
  var formId = s[0];
  var recordId = s[1];
  var currentRecordId = getRecordId(formId);
  var detailModule = getFormDetailModule(formId);
	if (currentRecordId != recordId)
	{
    var elem = document.getElementById(formId + gridRowIdSeparator + recordId);
  	elem.className += " selRow";
  	if (currentRecordId >= 0)
  	{
      elem = document.getElementById(formId + gridRowIdSeparator + currentRecordId);
      if (elem)
      {
        var i = elem.className.indexOf("selRow");
        if (i <= 1)
          elem.className = "";
        else
          elem.className = elem.className.substring(0, i - 1);
      }
  	}
    setRecordId(formId, recordId);
  	ajaxAbort();
    if (detailModule)
      loadDetailModule(detailModule, recordId);
    else
      ajaxGetText(buildRequestUrl("recordView", formId) + "&selectedByUser=1", formId + "RecordView", loadingSignalId, 
        function() {
        setLiteralKeyValue(formId);
        loadAllDescendantForms(formId);
        }
      );
	}
  else if (detailModule)
    loadDetailModule(detailModule, recordId);
}

function scrollSelectedItemIntoView(formId, immediate)
{
  var recordId = getRecordId(formId);
  if (!immediate)
  {
  	if (recordId >= 0)
      setTimeout(function() { scrollSelectedItemIntoView(formId, true) }, 1000); // give time to load images into grid.
  }
  else
  {
	  var elem = document.getElementById(formId + gridRowIdSeparator + recordId);
    var gridElem = document.getElementById(formId + "Grid");
    if (elem && gridElem)
    {
      if ((elem.offsetHeight >= gridElem.offsetHeight) || (elem.offsetTop <= gridElem.scrollTop))
        gridElem.scrollTop = elem.offsetTop;
      else if (elem.offsetTop + elem.offsetHeight >= gridElem.scrollTop + gridElem.offsetHeight)
        gridElem.scrollTop = elem.offsetTop + elem.offsetHeight - gridElem.offsetHeight + 2; // 2px = grid border.
    }
	}
}

function refreshForm(formId)
{
  if (inserting && (insertingFormId == formId))
  {
    inserting = false;
    insertingFormId = "";
  }
	loadingForms = true;
  ajaxGetText(buildRequestUrl("gridView", formId), formId + "GridView", loadingSignalId, 
    function(){
      scrollSelectedItemIntoView(formId);
      if (formCheckCurrentRecordVisible == formId && !formCurrentRecordIsVisible(formId))
      {
        formCheckCurrentRecordVisible = null;
        delay(100,
          function() {
            alert("El registro con el que estaba trabajando quedó fuera de la consulta actual.");
          }
        );
      }
      ajaxGetText(buildRequestUrl("recordView", formId), formId + "RecordView", loadingSignalId, 
        function() {
          setLiteralKeyValue(formId);
          if (!inserting)
            loadAllDescendantForms(formId);
          else
            loadingForms = false;
        }
      );
  	}
  );
}

function refreshSingleForm(formId, onFinished)
{
  if (inserting && (insertingFormId == formId))
  {
    inserting = false;
    insertingFormId = "";
  }
	ajaxGetText(buildRequestUrl("gridView", formId), formId + "GridView", loadingSignalId, 
    function(){
      scrollSelectedItemIntoView(formId);
      ajaxGetText(buildRequestUrl("recordView", formId), formId + "RecordView", loadingSignalId, 
        function() {
          setLiteralKeyValue(formId);
          if (onFinished)
            if (typeof(onFinished) == "function")
              onFinished()
            else
              eval(onFinished);
        }
      );
  	}
  );
}

function refreshDependantForms(formId)
{
  var dependantFormIds = getDependantFormIds(formId);
	if (dependantFormIds != "none")
    for (var i = 0; i < dependantFormIds.length; i++)
      refreshForm(dependantFormIds[i]);
}

function refreshParentForm(formId)
{
  var pId = getParentFormId(formId);
  if (pId != "none")
  {
    if (parentFormIsDependant(formId))
    {
      ajaxSynchronousGetText(buildRequestUrl("gridView", pId), pId + "GridView", loadingSignalId);
      scrollSelectedItemIntoView(pId);
      if (formCheckCurrentRecordVisible == formId && !formCurrentRecordIsVisible(pId))
      {
        formCheckCurrentRecordVisible = null;
        delay(100,
          function() {
            alert("El registro con el que estaba trabajando quedó fuera de la consulta actual.");
          }
        );
      }
      ajaxSynchronousGetText(buildRequestUrl("recordView", pId), pId + "RecordView", loadingSignalId); 
      refreshParentForm(pId);
    }
  }
}

var formIds = null; // array of childIds.
var formLevelIdx = null; // array of level indexes
var formIdList = null; // flat list of formIds.

function loadFormContainer(formContainerId, containerId)
{
  formIds = new Array;
  formLevelIdx = new Array;
  formIdList = new Array;
  currentFormContainerId = formContainerId;
  document.getElementById(containerId).innerHTML = "";
  ajaxGetText(server + "?sessionId=" + sessionId + "&verb=formContainer&formId=" + formContainerId, containerId, loadingSignalId,
    function() {
      var e = document.getElementById(formContainerId + "MainPanel");
      e.scrollLeft = 0;
      e.scrollTop = 0;
      document.title = document.getElementById("formContainerTitleText").innerHTML + " - " + appBaseWindowTitle;
      loadAllDescendantForms(formContainerId);
    }
  );
}

function findNextFormId()
{
  var found = false;
  do
  {
    var i = formIds.length - 1;
    var childFormIds = getChildFormIds(formIds[i][formLevelIdx[i]]);
    if (childFormIds != "none")
    {
      i++;
      formIds[i] = childFormIds;
      formLevelIdx[i] = 0;
      found = true;
    }
    else
    {
      do
      {
        formLevelIdx[i]++;
        if (formLevelIdx[i] == formIds[i].length)
        {
          formIds.splice(i, 1);
          formLevelIdx.splice(i, 1);
          i--;
        }
        else
          found = true;
      } while (!found && formIds.length > 0);
    }
  } while (!found && formIds.length > 0);
  return(found);
}

var refreshStartingFormId = "";
var refreshCount = 0;

function loadNextForm()
{
  if (findNextFormId())
  {
    var i = formIds.length - 1;
    loadForm(formIds[i][formLevelIdx[i]]);
  }
  else 
  {
    if (refreshStartingFormId)
    {
      refreshDependantForms(refreshStartingFormId);
      refreshStartingFormId = "";
    }
    else
      showAllForms();
    loadingForms = false;
  }
}

function loadForm(formId)
{
  if (document.getElementById(formId).innerHTML)
  {
    setRecordId(formId, -1);
    document.getElementById(formId + "Grid").innerHTML = "";
    ajaxGetText(buildRequestUrl("gridView", formId), formId + "GridView", loadingSignalId, 
      function(){
        scrollSelectedItemIntoView(formId);
        ajaxGetText(buildRequestUrl("recordView", formId), formId + "RecordView", loadingSignalId, 
          function() {
            setLiteralKeyValue(formId);
            loadNextForm();
          }
        );
      }
    );
  }
  else
    ajaxGetText(buildRequestUrl("form", formId), formId, loadingSignalId, 
      function(){
        formIdList[formIdList.length] = formId;
        ajaxGetText(buildRequestUrl("gridView", formId), formId + "GridView", loadingSignalId, 
          function(){
            scrollSelectedItemIntoView(formId);
            ajaxGetText(buildRequestUrl("recordView", formId), formId + "RecordView", loadingSignalId, 
              function() {
                setLiteralKeyValue(formId);
                loadNextForm();
              }
            );
          }
        );
      }
    );
}

var loadingForms = false;

function loadAllDescendantForms(formId)
{
  loadingForms = true;
  formIds = new Array;
  formIds[0] = new Array;
  formIds[0][0] = formId;
  formLevelIdx[0] = 0;
  loadNextForm();
}

function showAllForms()
{
  for (var i = 0; i < formIdList.length; i++)
    document.getElementById(formIdList[i]).style.visibility = "visible";
  if (firstFormRecordId >= 0 && firstFormRecordId != getRecordId(formIdList[0]))
    alert("El registro seleccionado no está visible en la consulta actual.");
  firstFormRecordId = -1;
}

function refreshAllForms(formId)
{
  refreshStartingFormId = formId;
  refreshForm(formId);
}

var inserting = false;
var insertingFormId = "";

function formSetFocus(formId)
{
  var frmElems = document.forms[formId].elements;
  for (var i = 0; i < frmElems.length; i++)
  {
    var elem = frmElems[i];
    if (!elem.readOnly && (elem.className.toLowerCase() != "hidden"))
    {
      elem.focus();
      return;
    }
  }
}

function formElementKeyUp(evt) {
  evt = (evt) ? evt : ((window.event) ? event : null);
  if (evt)
  {
    var elem = (evt.target) ? evt.target : ((evt.srcElement) ? evt.srcElement : null);
    if (elem && elem.form)
    {
      if (evt.keyCode == 13 && evt.ctrlKey)
        formSave(elem.form.name);
      else if (evt.keyCode == 27)
      {
        inserting = false;
        insertingFormId = "";
        refreshForm(elem.form.name);
      }
    }
  }
}

function formAddNew(formId)
{
  if (loadingForms) return;
  var parentFormId = getParentFormId(formId);
  if (parentFormId == "none" || (getRecordId(parentFormId) >= 0))
	{
	  deselectGridViewItem(formId);
    setRecordId(formId, -1);
    ajaxGetText(buildRequestUrl("newRecordView", formId), formId + "RecordView", loadingSignalId, 
      function() {
        inserting = true;
        insertingFormId = formId;
        setLiteralKeyValue(formId);
        loadAllDescendantForms(formId);
        formSetFocus(formId);
      }
    );
	}
}

function formDelete(formId)
{
  if (loadingForms) return;
  var recordId = getRecordId(formId);
  if (recordId >= 0)
	{
    if (!confirm("Confirme si desea eliminar el elemento seleccionado.")) return;
  	var childFormIds = getChildFormIds(formId);
  	if (childFormIds != "none" && !cascadeDeleteEnabled(formId) && !formsEmpty(childFormIds))
    {
      alert("No es posible borrar el elemento seleccionado porque exiten datos que dependen de él.");
      return;
    }
  	if (childFormIds == "none" || formsEmpty(childFormIds) ||
        confirm("El elemento que desea eliminar tiene otros elementos dependientes que tambien seran eliminados.\n\nConfirma la operacion?"))
      ajaxGetText(buildRequestUrl("delete", formId), "", loadingSignalId, 
  			function() {
          if (ajaxResponseText)
          {
    			  eval("var obj = " + ajaxResponseText);
    			  if (obj.result == "OK")
            {
              setRecordId(formId, -1);
              refreshParentForm(formId);
    					refreshAllForms(formId);
            }
            if (obj.message)
    				  alert(obj.message);
          }
  			}
			);
	}
}

function formMove(formId, direction)
{
  if (loadingForms) return;
	var recordId = getRecordId(formId);
	if (recordId >= 0)
  	ajaxGetText(buildRequestUrl("move", formId) + "&direction=" + direction, formId + "GridView", loadingSignalId, 
  		function() {
        if (ajaxResponseText)
        {
  			  eval("var obj = " + ajaxResponseText);
  			  if (obj.result == "OK")
          {
            setRecordId(formId, obj.recordId);
      			ajaxGetText(buildRequestUrl("gridView", formId), formId + "GridView", loadingSignalId);
          }
  				if (obj.message)
  				  alert(obj.message);
        }
  		}
		);
}

var formCheckCurrentRecordVisible = null;

function formSave(formId)
{
  if (loadingForms) return;
  var parentFormId = getParentFormId(formId);
  if (((parentFormId == "none") || (getRecordId(parentFormId) >= 0)) && beforePost(formId))
	{
    ajaxPostText(buildRequestUrl("update", formId), "", serializeForm(document.forms[formId]), loadingSignalId,
			function () {
			  try { afterPost(formId) } catch (e) { alert("Error en afterPost(" + formId + ")") }
        if (ajaxResponseText)
        {
  			  eval("var obj = " + ajaxResponseText);
  			  if (obj.result == "OK")
          {
            if (inserting && (insertingFormId == formId) && formAutoInsert(formId))
              delay(2000,
                function () {
                  formAddNew(formId);
                }
              );
            setRecordId(formId, obj.recordId);
            refreshParentForm(formId);
            formCheckCurrentRecordVisible = formId;
            refreshAllForms(formId);
          }
  				if (obj.message)
  				  alert(obj.message);
        }
			}
		);
	}
}

var uploadingFile = false;
var uploadingFileDataType = "archivo";
var uploadingMaxFilesize = 0;

function fileUploadDialog(formId, dbFieldBaseName, fileDataTypeName, recomendedMaxFilesize)
{
  if (loadingForms) return;
  if (fileDataTypeName)
    uploadingFileDataType = fileDataTypeName;
  else
    uploadingFileDataType = "archivo";
  uploadingMaxFilesize = recomendedMaxFilesize;
	var recordId = getRecordId(formId);
	if (recordId >= 0)
  	ajaxGetText(buildRequestUrl("fileUploadDialog", formId) + "&dbFieldBaseName=" + dbFieldBaseName + "&dataType=" + uploadingFileDataType, 
		  "dialogsPanel", loadingSignalId,
  	  function () {
  		  showDialog("fileUploadDialog");
  		}
   	);
	else
	  alert("Solo se puede cargar contenido (" + fileDataTypeName.toLowerCase() + 
		  ") en un elemento existente (guardado).\n\nSeleccione uno de la lista y vuelva a intentarlo.");
}

function fileUploadOK(formId)
{
  var frm = document.formFileUpload;
  if (!uploadingFile && (frm.uploadFormFileField.value || 
	  confirm("No ha elegido un archivo.\n\n¿Desea eliminar el contenido (" + uploadingFileDataType + ") existente?")))
	{
    if (!frm.uploadFormFileField.files || frm.uploadFormFileField.files.length == 0 ||
         frm.uploadFormFileField.files[0].size <= uploadingMaxFilesize ||
         confirm("El tamaño del archivo es de " + Math.round(frm.uploadFormFileField.files[0].size / 1024) + " KB, " +
           "y excede el máximo recomendado de " + Math.round(uploadingMaxFilesize / 1024) + " KB." +
           "\n\n¿Desea cargarlo de todos modos?"))
    {
      uploadingFile = true;
      document.getElementById(loadingSignalId).style.visibility = "visible";
      document.formFileUpload.submit();
    }
	}
	else
	  fileUploadCancel(formId);
}

function fileUploadCancel(formId)
{
  if (!uploadingFile)
		hideDialog();
}

function fileUploaded(formId)
{
  if (uploadingFile)
	{
		uploadingFile = false;
		hideDialog();
		ajaxGetText(buildRequestUrl("gridView", formId), formId + "GridView", loadingSignalId, 
		  function() {
        scrollSelectedItemIntoView(formId);
        ajaxGetText(buildRequestUrl("recordView", formId), formId + "RecordView", loadingSignalId);
			}
		);
	}
}

var htmlContentFieldFormId = "";
var htmlContentFieldName = "";
var htmlFieldElem = null;

function htmlEditorDialog(formId, fieldName, fieldLabel)
{
  if (loadingForms) return;
  htmlFieldElem = document.forms[formId].elements[fieldName];
  try
  {
    var editor = FCKeditorAPI.GetInstance('FCKeditor1');
    editor.SetData(htmlFieldElem.value, true);
    document.getElementById("htmlEditorTitle").innerHTML = fieldLabel;
    showDialog("htmlEditorDialog");
  }
  catch (e)
  {
  	ajaxGetText(buildRequestUrl("htmlEditorDialog", formId), "dialogsPanel", loadingSignalId,
  	  function ()	{
        editor = new FCKeditor("FCKeditor1", "100%", "100%", "Default", htmlFieldElem.value);
        editor.BasePath = document.location.pathname.substring(0,document.location.pathname.lastIndexOf('back')) + "fckeditor/" ;
        document.getElementById("htmlEditorDiv").innerHTML = editor.CreateHtml();
        document.getElementById("htmlEditorTitle").innerHTML = fieldLabel;
        showDialog("htmlEditorDialog");
      }
   	);
  }
}

function htmlEditorOK()
{
  var editor = FCKeditorAPI.GetInstance('FCKeditor1');
  var editorContent = editor.GetXHTML(false);
  if (editorContent != htmlFieldElem.value)
		htmlFieldElem.value = editorContent;
	hideDialog();
  editor.SetData("", true);
}

function htmlEditorCancel()
{
  var editor = FCKeditorAPI.GetInstance('FCKeditor1');
  if (editor.GetXHTML(false) == htmlFieldElem.value ||
      confirm("Ha realizado cambios al contenido. Confirme si realmente desea descartarlos."))
  {
  	hideDialog();
    editor.SetData("", true);
  }
}

var validating = false;
var validatingCtrl = null;

function checkNull(formId, fieldName)
{
  if (validating)
	{
  	validatingCtrl = document.forms[formId].elements[fieldName + "NewValue"];
  	if (!validatingCtrl.value)
  	{
  	  alert("El campo " + document.getElementById(formId + fieldName + "Label").innerHTML + " es obligatorio.");
  		setTimeout(function(){validatingCtrl.focus()}, 100);
  	}
  	else
    	validatingCtrl = null;
	}
}

function normalPost(formId, action, target)
{
	var frm = document.forms[formId];
  frm.action = action;
  frm.target = target;
  frm.method = "post";
  if (beforePost(formId))
    frm.submit();
}

function beforePost(formId)
{
  validating = true;
	validatingCtrl = null;
	var frm = document.forms[formId];
	for (var i = 0; !validatingCtrl && (i < frm.elements.length); i++)
  	if (frm.elements[i].onblur)
  	  frm.elements[i].onblur();
  validating = false;
  return(!validatingCtrl);
}

function afterPost(formId)
{
  return(true);
}

var lastSearchFieldName = "";
var lastSearchFieldValue = "";

function lookupSearchFieldChanged(formId, elem)
{
  if (loadingForms)
  {
    elem.value = lastSearchFieldValue;
    return;
  }
  if (timeoutHandle)
    clearTimeout(timeoutHandle);
  if (lastSearchFieldName != elem.name)
  {
    lastSearchFieldName = elem.name;
    lastSearchFieldValue = "";
  }
  if (lastSearchFieldValue != elem.value)
  {
    lastSearchFieldValue = elem.value;
    timeoutHandle = setTimeout(
      function() {
        timeoutHandle = null;
        var frm = document.forms[formId];
        var s = lastSearchFieldName.replace("SearchValue", "");
        frm.elements[s + "LookupList"].innerHTML = "<option selected=selected>cargando...</option>";
        ajaxGetText(buildRequestUrl("getLookupSearchFieldOptions", formId) +
          "&lookupTable=" + encodeURIComponent(frm.elements[s + "LookupTable"].value) +
          "&lookupOrder=" + encodeURIComponent(frm.elements[s + "LookupOrder"].value) +
          "&columnExpr=" + encodeURIComponent(frm.elements[s + "ColumnExpr"].value) +
          "&searchValue=" + encodeURIComponent(lastSearchFieldValue), "", loadingSignalId,
          function() {
            frm.elements[s + "LookupList"].innerHTML = ajaxResponseText;
          }
        );
      }, 2000
    );
  }
}

function lookupSearchValueChanged(formId, elem)
{
  var frm = document.forms[formId];
  var s = elem.name.replace("LookupList", "");
  frm.elements[s + "ReadOnly"].value = elem.options[elem.selectedIndex].text;
  frm.elements[s + "NewValue"].value = elem.value;
}

// MISCELANEOUS FUNCTIONS =============================================================================================

var delayTask = null;

function delay(time, task)
{
  delayTask = task;
  setTimeout(
	  function() {
		  if (typeof(delayTask) == "function")
			  delayTask();
			else
			  eval(delayTask);
		},
    time
  );
}



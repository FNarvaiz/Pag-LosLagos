
var appServer = "front-end/main.asp";
var loadingSignalId = "loadingSignal";
var lang = "SP";
var sessionId = "";
var usrName = "";
var contador;
function track(category, action, label)
{
  try
  {
    _gaq.push(['_trackEvent', category, action, label]);
  }
  catch(e) {}
}
function menu () {
  if(contador!=undefined){
    if (contador == 1) {
      var elem=document.getElementById("navegador");
      if(elem!=null)
        elem.style.left="0px";
      else
        document.getElementById('loginMenuItem').click();
      contador = 0;
    } else {
      contador = 1;
      document.getElementById("navegador").style.left="-100%";
      
    }
  }
}
function bookingsSendRequest2(bookingDate,placeId, turnId)
{
  if (!bookingDate || !turnId) {
    alert("Por favor, completá todos los datos para poder solicitar la reserva.");
    return;
  }
  ajaxGetText(appServer + "?lang=" + lang + "&sessionId=" + sessionId + "&content=bookingsSendRequest" + 
    "&bookingDate=" + encodeURIComponent(bookingDate) + "&placeId=" + placeId + "&turnId=" + turnId, "", loadingSignalId,
    function() {
      eval("var response = " + ajaxResponseText);
      if (response.message) {
        alert(response.message);
        document.getElementById("bookingsFormDate").value = "";
        document.getElementById("bookingsFormTurn").selectedIndex = 0;
      }
    }
  );
}
function placeMainElements()
{
  var elem = document.getElementById("mainPanel");
  if (elem.offsetParent.clientWidth - elem.offsetWidth > 0)
    elem.style.left = Math.round((elem.offsetParent.clientWidth - elem.offsetWidth) / 2) + "px";
  if (elem.offsetParent.clientHeight - elem.offsetHeight > 0)
  	elem.style.top = Math.round((elem.offsetParent.clientHeight - elem.offsetHeight) / 2) + "px";
  document.getElementById(loadingSignalId).style.left = (elem.offsetLeft + elem.offsetWidth - 40) + "px";
  document.getElementById(loadingSignalId).style.top = (elem.offsetTop + 6) + "px";
}

function init(langCode, usr, pwd)
{
  if (langCode)
    lang = langCode;
  bodyResized();
  showLoginDialog();
}
function showLoginDialog()
{
  ajaxGetText(appServer + "?lang=" + lang + "&content=loginDialog", "dynPanel", "",
    function(){
      document.getElementById("loginDialog").style.display = "block";
      ajaxGetText(appServer + "?lang=" + lang + "&content=mainMenu", "mainMenu", "",
        function()
        {
          mainMenuOptionChanged(document.getElementById("loginMenuItem"));
          document.loginForm.usr.focus();
        }
      );
    }
  );
}
function done()
{
  if (sessionId)
  {
    track("sesion", "cerrar", "Familia " + usrName);
    ajaxSynchronousGetText(appServer + "?lang=" + lang + "&sessionId=" + sessionId + "&content=logout", "");
  }
  else
    track("sesion", "cerrar", "sin login");
  sessionId = "";
}



function login(usr, pwd)
{
  if (!usr || !pwd)
  {
    alert("Por favor, ingrese los datos de acceso.");
    return;
  }
  ajaxLoadingSignalOn(loadingSignalId);
  ajaxGetText(appServer + "?lang=" + lang + "&content=login&usr=" + usr + "&pwd=" + pwd, "", "",
    function() {
      eval("var response = " + ajaxResponseText);
      if (response.result == "failed")
      {
        track("sesion", "login fallido", "usr: " + usr + "- pwd: " + pwd);
        ajaxLoadingSignalOff();
        alert(response.message);
        showLoginDialog();
      }
      else
      {
        document.getElementById("dynPanel").style.backgroundColor= "#badbe0";
        sessionId = response.sessionId;
        usrName = response.usrName;
        track("sesion", "login OK", "Familia " + usrName);
        start();
      }
    }
  );
}

function start()
{
  ajaxGetText(appServer + "?lang=" + lang + "&sessionId=" + sessionId + "&content=mainMenu", "mainMenu", "",
    function() {
      ajaxGetText(appServer + "?lang=" + lang + "&sessionId=" + sessionId + "&content=userMenu", "userMenu", "",
        function() {
          document.getElementById("mainMenuOption4").click();
        }
      );
    }
  );
}

function logout()
{
  location.reload();
}

var prevMainMenuSelectedOption = null;

function showChangePasswordDialog()
{
  if (ajaxBusy()) return;
  prevMainMenuSelectedOption = mainMenuSelectedOption;
  mainMenuOptionChanged(null);
  ajaxGetText(appServer + "?lang=" + lang + "&content=changePasswordDialog", "dynPanel", "",
    function(){
      document.changePasswordForm.pwd.focus();
    }
  );
}

function hideChangePasswordDialog()
{
  if (ajaxBusy()) return;
  if (prevMainMenuSelectedOption) prevMainMenuSelectedOption.click();
}

function changePassword()
{
  var currentPwd = document.changePasswordForm.pwd.value;
  var pwd1 = document.changePasswordForm.pwd1.value;
  var pwd2 = document.changePasswordForm.pwd2.value;
  if (currentPwd && pwd1 && pwd2)
  {
    if (pwd1 == pwd2)
    {
      ajaxLoadingSignalOn(loadingSignalId);
      ajaxGetText(appServer + "?lang=" + lang + "&sessionId=" + sessionId + 
          "&content=changePassword&currentPwd=" + currentPwd + "&newPwd=" + pwd1, "dynPanel", "",
        function() {
          track("sesion", "cambio clave", "Familia " + usrName);
          ajaxLoadingSignalOff();
        }
      );
    }
    else
    {
      if (lang == "SP")
        alert("La nueva contraseña y su repetición deben ser idénticas.");
      else
        alert("The new and the repeated passwords must be identical.");
    }
  }
  else
  {
    if (lang == "SP")
      alert("Por favor, completá todos los datos.");
    else
      alert("Please, complete every field.");
  }
}

function showDataViewer(contentId)
{
  ajaxGetText(appServer + "?lang=" + lang + "&sessionId=" + sessionId + "&content=textContent&contentId=" + contentId, 
    "dataViewerPanel", loadingSignalId,
    function() {
      fade("dataViewer", true);
    }
  );
}

function hideDataViewer()
{
  fade("dataViewer", false,
    function() {
      document.getElementById("dataViewerPanel").innerHTML = "";
      }
    );
}

function showImgViewer(imgElem)
{
  setupImg(imgElem, "imgViewerPanel", "imgViewerImage", 3)
//  document.getElementById("imgViewerImage").src = imgElem.src;
  fade("imgViewer", true);
}

function hideImgViewer()
{
  fade("imgViewer", false,
    function() {
      document.getElementById("imgViewerImage").style.visibility = "hidden";
      document.getElementById("imgViewerImage").src = "";
      }
    );
}

function saveSurveyVote(confirmationQuestion, surveyId, voteValue, itemNumber)
{
  if (!confirmationQuestion || confirm(confirmationQuestion))
  {
    ajaxGetText(appServer + "?lang=" + lang + "&sessionId=" + sessionId + "&content=saveSurveyVote&surveyId=" + surveyId + 
      "&voteValue=" + voteValue + "&itemNumber=" + itemNumber, "surveysItem" + itemNumber, loadingSignalId,
      function()
      {
        var e = document.getElementById("voteComment" + itemNumber);
        e.style.height = "70px";
        e.style.display = "block";
        track("encuestas", "valoracion", "Familia " + usrName);
      }
    );
  }
}

function saveSurveyVoteComment(confirmationQuestion, surveyId, itemNumber)
{
  if (!confirmationQuestion || confirm(confirmationQuestion))
  {
    var voteComment = document.getElementById("voteCommentText" + itemNumber).value;
    ajaxGetText(appServer + "?lang=" + lang + "&sessionId=" + sessionId + "&content=saveSurveyVoteComment&surveyId=" + surveyId + 
      "&voteComment=" + encodeURIComponent(voteComment) + "&itemNumber=" + itemNumber, "surveysItem" + itemNumber, loadingSignalId,
      function()
      {
        track("encuestas", "Comentario", "Familia " + usrName);
      }
    );
  }
}

function saveSupplierVote(confirmationQuestion, supplierId, voteValue, itemNumber)
{
  if (!confirmationQuestion || confirm(confirmationQuestion))
  {
    ajaxGetText(appServer + "?lang=" + lang + "&sessionId=" + sessionId + "&content=saveSupplierVote&supplierId=" + supplierId + 
      "&voteValue=" + voteValue + "&itemNumber=" + itemNumber, "suppliersItem" + itemNumber, loadingSignalId,
      function()
      {
        track("proveedores", "valoracion", "Familia " + usrName);
      }
    );
  }
}

function downloadCategorySelected(combobox)
{
  document.getElementById("downloadsListing").innerHTML = "";
  document.getElementById("downloadsListing").scrollTop = 0;
  ajaxGetText(appServer + "?lang=" + lang + "&sessionId=" + sessionId + "&content=downloadListing&categoryId=" + combobox.value, "downloadsListing", loadingSignalId,
    function()
    {
      track("historico-mostrar", combobox.options[combobox.selectedIndex].text, "Familia " + usrName);
    }
  );
}

function suppliersCategorySelected(combobox)
{
  document.getElementById("suppliersPanel").innerHTML = "";
  document.getElementById("suppliersPanel").scrollTop = 0;
  ajaxGetText(appServer + "?lang=" + lang + "&sessionId=" + sessionId + "&content=suppliersListing&categoryId=" + combobox.value, "suppliersPanel", loadingSignalId,
    function()
    {
      track("proveedores-rubro", combobox.options[combobox.selectedIndex].text, "Familia " + usrName);
    }
  );
}

function classifiedsCategorySelected(combobox)
{
  document.getElementById("classifiedsPanel").innerHTML = "";
  document.getElementById("classifiedsPanel").scrollTop = 0;
  ajaxGetText(appServer + "?lang=" + lang + "&sessionId=" + sessionId + "&content=classifiedsListing&categoryId=" + combobox.value, "classifiedsPanel", loadingSignalId,
    function()
    {
      prepareTwoColumnsLayout("classifieds");
      track("avisos-rubro", combobox.options[combobox.selectedIndex].text, "Familia " + usrName);
    }
  );
}
function bookingResourceSelectedHOME()
{
  if(document.getElementById("bookingsFormPlace") != undefined){
    var placeId = document.getElementById("bookingsFormPlace").value;
    if(placeId==0){
          bookingResourceFieldChanged(30);
    }
    else {
      if(placeId==1){
          bookingResourceFieldChanged(40);
      }
      else{
        bookingResourceFieldChanged(29);
      }
    }
  }
}
function comboBoxAppendOption(combobox, text, value)
{
  var elOptNew = document.createElement('option');

  elOptNew.text = text;
  elOptNew.value = value;
  try {
    combobox.add(elOptNew, null); // standards compliant; doesn't work in IE
  }
  catch(e) {
    combobox.add(elOptNew); // IE only
  }
}
function bookingResourceFieldChanged(idrecurso)
{
  var date=document.getElementById("bookingsFormDate").value;
  var bookingStartComboBox = document.getElementById("bookingsFormTurn");
  
  
  ajaxGetText(appServer + "?lang=" + lang + "&sessionId=" + sessionId + "&content=jBookingStartOptions&resourceId=" + idrecurso+"&bookingDate="+encodeURIComponent(date),"", loadingSignalId,
    function() {
      try{
        eval("var response = " + ajaxResponseText);
        bookingStartComboBox.length = 0;
        for (i in response.bookingStartOptions)
          comboBoxAppendOption(bookingStartComboBox, response.bookingStartOptions[i].name, response.bookingStartOptions[i].id);
        bookingStartComboBox.selectedIndex = 0;
      } catch(e) {
        bookingStartComboBox.length=0;
        comboBoxAppendOption(bookingStartComboBox, "No disponible", "0");
      }
    });

}
function datapicker_onclick() {
   numberOfClicks++;
   window.document.form1.myButton.value = 'Botón pulsado ' + numberOfClicks + 
   ' veces';
}
function bookingResourceSelected(combobox)
{
  document.getElementById("bookingsStatus").innerHTML = "";
  document.getElementById("bookingsStatus").scrollTop = 0;
  ajaxGetText(appServer + "?lang=" + lang + "&sessionId=" + sessionId + "&content=bookingsStatus&resourceId=" + combobox.value, "bookingsStatus", loadingSignalId,
    function()
    {
      track("reservas-estado", combobox.options[combobox.selectedIndex].text, "Familia " + usrName);
    }
  );
}

function bookingsSendRequest()
{
  var bookingDate = document.getElementById("bookingsFormDate").value;
  var placeId = document.getElementById("bookingsFormPlace").value;
  var turnId = document.getElementById("bookingsFormTurn").value;
  if (!bookingDate || !placeId || !turnId) {
    alert("Por favor, indicá la Fecha del evento, el Espacio y el Turno para poder solicitar la reserva.");
    return;
  }
  ajaxGetText(appServer + "?lang=" + lang + "&sessionId=" + sessionId + "&content=bookingsSendRequest" + 
    "&bookingDate=" + encodeURIComponent(bookingDate) + "&placeId=" + placeId + "&turnId=" + turnId, "", loadingSignalId,
    function() {
      eval("var response = " + ajaxResponseText);
      if (response.message) {
        alert(response.message);
        document.getElementById("bookingsFormDate").value = "";
        document.getElementById("bookingsFormTurn").selectedIndex = 0;
      }
    }
  );
}

function bookingsCancel(bookingId, resourceId)
{
  if (confirm("¿Confirma que desea cancelar la reserva?"))
    ajaxGetText(appServer + "?lang=" + lang + "&sessionId=" + sessionId + "&content=bookingCancel&bookingId=" + bookingId, "", loadingSignalId,
      function() {
        eval("var response = " + ajaxResponseText);
        if (response.result == "failed")
          alert(response.message);
        else
        {
          var i = document.getElementById("bookingRows").scrollTop;
          ajaxGetText(appServer + "?lang=" + lang + "&sessionId=" + sessionId + "&content=bookingsStatus&resourceId=" + resourceId,
            "bookingsStatus", loadingSignalId,
            function() {
              var bookingsResourceSelector = document.getElementById("bookingsResourceSelector");
              for (var j = 0; j < bookingsResourceSelector.options.length; j++)
                if (bookingsResourceSelector.options[j].value == resourceId) {
                  bookingsResourceSelector.selectedIndex = j;
                  break;
                }
              document.getElementById("bookingRows").scrollTop = i;
              ajaxGetText(appServer + "?lang=" + lang + "&sessionId=" + sessionId + "&content=bookingsListing", "bookingsListing", loadingSignalId);
            }
          );
        }
      }
    );
}

function bookingsTake(elem, resourceId, bookingDate, turnStart, turnType, requiresExtraNeighborUnit)
{
  var extraNeighborUnit = "";
  if (requiresExtraNeighborUnit) {
    extraNeighborUnit = window.prompt("Nro. de lote del vecino con quien comparte la reserva:", "");
    if (!extraNeighborUnit) return;
  }
  ajaxGetText(appServer + "?lang=" + lang + "&sessionId=" + sessionId + "&content=bookingsTake" +
    "&resourceId=" + resourceId + "&bookingDate=" + encodeURIComponent(bookingDate) + "&turnStart=" + turnStart + 
    "&turnType=" + turnType + "&extraNeighborUnit=" + extraNeighborUnit, "", loadingSignalId,
    function() {
      eval("var response = " + ajaxResponseText);
      if (response.result == "failed")
        alert(response.message);
      else
      {
        var i = document.getElementById("bookingRows").scrollTop;
        ajaxGetText(appServer + "?lang=" + lang + "&sessionId=" + sessionId + "&content=bookingsStatus&resourceId=" + resourceId,
          "bookingsStatus", loadingSignalId,
          function() {
             document.getElementById("bookingRows").scrollTop = i;
             ajaxGetText(appServer + "?lang=" + lang + "&sessionId=" + sessionId + "&content=bookingsListing", "bookingsListing", loadingSignalId);
          }
        );
      }
    }
  );
}

// Menu functions

var mainMenuSelectedOption = null;

function mainMenuOptionChanged(elem)
{
  if (mainMenuSelectedOption)
  {
    mainMenuSelectedOption.style.backgroundImage = "";
    mainMenuSelectedOption.style.color = "";
    mainMenuSelectedOption.style.borderLeft="";
  }
	mainMenuSelectedOption = elem;
  if (mainMenuSelectedOption)
  {
    mainMenuSelectedOption.style.backgroundImage = "url(front-end/resource/menuBtnBgSelected.png)";
    mainMenuSelectedOption.style.color = "#404040";
    mainMenuSelectedOption.style.borderLeft="4px solid #31cccc";
  }
}

function isJSON(text)
{
  return(typeof text === "string" && text.length > 0 && text[0] == "{" && text[text.length - 1] == "}");
}

function load(menuElement, contentId)
{
  contador = 1;
  if (document.body.clientWidth<900) {
   
   if(contentId!="faq" && contentId!="registrationDialog" && contentId!="passwordRecoveryDialog" && contentId!="loginMenuItem")
      document.getElementById("navegador").style.left="-100%";

  }

  if((contentId=="profile" || contentId=="timeLine") && document.body.clientWidth<900){
    document.getElementById("dynPanel").innerHTML="<div class='noView'>Esta sección se ve solo en pantalla grande para un mejor funcionamiento.</div>";
  }
  else{
    mainMenuOptionChanged(menuElement);
    ajaxGetText(appServer + "?lang=" + lang + "&sessionId=" + sessionId + "&content=" + contentId, "", loadingSignalId,
      function() {
        if (isJSON(ajaxResponseText))
        {
          eval("var response = (" + ajaxResponseText + ")");
          if (response.result == "failed")
          {
            alert(response.message);
            showLoginDialog();
          }
        }
        else
        {
          track("seccion", (!menuElement ? "" : menuElement.innerHTML) , "Familia " + usrName);
          document.getElementById("dynPanel").innerHTML = ajaxResponseText;
          switch (contentId)
          {
            case "downloads":
              document.getElementById("downloadsCategorySelector").selectedIndex = 0;
              document.getElementById("downloadsCategorySelector").onchange();
              break;
            case "timeLine":
              var container = document.getElementById("timeLineGridContainer");
              if (container.scrollWidth <= container.offsetWidth)
                document.getElementById("timeLineArrows").style.visibility = "hidden";
              else
              {
                document.getElementById("timeLineArrows").style.visibility = "visible";
                container.scrollLeft = container.scrollWidth - container.offsetWidth;
              }
              timeLineColumnSelectedIdx = -1;
              timeLineDataTypeSelected = -1;
              eval(document.getElementById("javascriptCode").innerHTML);
              break;
            case "surveys":
              prepareTwoColumnsLayout("surveys");
              break;
            case "classifieds":
              document.getElementById("classifiedsCategorySelector").selectedIndex = 0;
              document.getElementById("classifiedsCategorySelector").onchange();
              break;
            case "profile":
              loadFormContainer("formAdminProfile", "dynPanel");
              refreshCurrentForm = false;
              break;
          }
        }
      }
    );
  }
}

function prepareTwoColumnsLayout(contentId)
{
  if (!contentId) return;
  var l = document.getElementById(contentId + "LeftCol");
  var r = document.getElementById(contentId + "RightCol");
  if (!l || !r) return;
  var elems = [];
  var i = 2;
  do {
    var e = document.getElementById(contentId + "Item" + i);
    if (e) 
      elems[i - 2] = e;
    i++;
  }
  while (e)
  for (i = 0; i < elems.length; i++)
    if (r.offsetHeight < elems[i].offsetTop) r.appendChild(elems[i]);
}

function sendFormData(formName)
{
  var frm = document.forms[formName];
  for (var i = 0; i < frm.elements.length; i++)
    if (!frm.elements[i].value)
    {
      if (lang == "SP")
        alert("Por favor, completá todos los datos del formulario.");
      else
        alert("Please, fill in all form fields.");
      return;
    }
  track("formulario", frm.trackingLabel.value, "Familia " + usrName);
  ajaxSubmit(frm, "", "loadingSignalId",
    function() {
      frm.reset();
      eval("var response = " + ajaxResponseText);
      if (response.message)
        alert(response.message);
    }
  );
}

function bodyResized()
{
  var main = document.getElementById("main");
  if (main)
  {
    if (document.body.clientWidth > main.offsetWidth)
      main.style.left = Math.round((document.body.clientWidth - main.offsetWidth) / 2) + "px";
    else
      main.style.left = "0px";
    if (document.body.clientHeight > main.offsetHeight)
      main.style.top = Math.round((document.body.clientHeight - main.offsetHeight) / 2) + "px";
    else
      main.style.top = "0px";
  }
}



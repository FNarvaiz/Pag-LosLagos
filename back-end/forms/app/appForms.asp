
<!--#include file="timeLineForms.asp"-->
<!--#include file="classifiedsForms.asp"-->
<!--#include file="reportsForms.asp"-->
<!--#include file="usersForms.asp"-->
<!--#include file="neighborsForms.asp"-->
<!--#include file="SurveysForms.asp"-->
<!--#include file="suppliersForms.asp"-->
<!--#include file="resourceBookingsForms.asp"-->
<!--#include file="contentsForms.asp"-->
<!--#include file="securityForms.asp"-->
<!--#include file="paramsForms.asp"-->

<%

' App messages

const noPermissionForOperation = "No tiene autorización para realizar esta operación."

const mainEmailName = "Vecinos de los Lagos"
const mainEmailAddress = "sistema@vecinosloslagos.com.ar"

' App render functions

function renderMain
  %>
  <div id="main">
    <div id="header">
      <div id="loadingSignal"></div>
    </div>
    <div id="topTitle">Sistema Brick
      <div id="menuBtn" style="display: none" onmouseover="showMenu()" onclick="menuBtnClick()">MENÚ</div>
      <div id="usrName">Usuario: <span id="usrNameLabel"></span> | <span class="anchor" onclick="window.close()">Cerrar</span> | 
        <span class="anchor" onclick="window.open(location.href)">Nueva Ventana</span></div>
    </div>
    <table id="initialMessage" width="100%" height="100%">
      <tr><td align="center" valign="middle"><img src="forms/app/resource/initializing.gif" onload="start()"><br>cargando...</td></tr>
    </table>
    <div id="menuBackgnd" style="display: none" onclick="hideMenu()">
      <div id="menu">
      </div>
    </div>
    <div id="mainPanel">
      <table id="welcomeMessage" width="100%" height="100%">
        <tr><td align="center" valign="middle">Para comenzar, seleccione una opción del menú.</td></tr>
      </table>
    </div>
    <div id="dialogbackgnd"></div>
    <div id="dialogsPanel"></div>
  </div>
  <%= renderDatePicker %>
  <%
end function

function renderMenuItem(itemName, formName)
  dim s
  if len(formName) = 0 then
    s = " style=" & dQuotes ("color: #666666")
  else
    s = "onclick=" & dQuotes("load(this, " & sQuotes(formName) & ")")
  end if
  %>
 	<div class="menuItem menuItemNormal" <%= s %> onmouseover="mouseOver(this)" onmouseout="mouseOut(this)"><%= itemName %></div>
  <%
end function

function renderMenuSeparator
  %>
 	<div class="menuSeparator"></div>
  <%
end function

function renderMenu
  if usrProfile = usrProfileSecurity then
    renderMenuItem "SEGURIDAD", "formAdminSecurity"
    exit function
  end if
  'renderMenuSeparator
  if usrAccessAdminNeighbors > usrPermissionNone then
    renderMenuItem "VECINOS", "formAdminNeighbors"
    renderMenuSeparator
  end if
  
  if usrAccessAdminTimeLine > usrPermissionNone then
    renderMenuItem "ARCHIVO HISTÓRICO", "formAdminTimeLine"
    renderMenuSeparator
  end if
  
  if usrAccessAdminBookings > usrPermissionNone then renderMenuItem "RESERVAS", "formAdminResourceBookings"

  if usrAccessAdminClassifieds > usrPermissionNone then renderMenuItem "AVISOS CLASIFICADOS", "formAdminClassifieds"
    
  if usrAccessAdminSurveys > usrPermissionNone then renderMenuItem "PROPUESTAS", "formAdminSurveys"

  if usrAccessAdminSuppliers > usrPermissionNone then renderMenuItem "PROVEEDORES", "formAdminSuppliers"

  if usrAccessAdminBookings > usrPermissionNone or usrAccessAdminClassifieds > usrPermissionNone or _
      usrAccessAdminSuppliers > usrPermissionNone or usrAccessAdminSuppliers > usrPermissionNone then
    renderMenuSeparator
  end if
  
  if usrAccessAdminUsers then
    renderMenuItem "USUARIOS ADMINISTRACIÓN", "formAdminUsers"
    renderMenuSeparator
  end if

  if usrAccessAdminReports > usrPermissionNone then
    renderMenuItem "INFORMES", "formAdminReports"
    renderMenuSeparator
  end if
  
 '' if usrAccessAdminContents > usrPermissionNone then
 ''   renderMenuItem "CONTENIDOS", "formAdminContents"
 ''   renderMenuSeparator
 '' end if

  if usrAccessAdminParams > usrPermissionNone then
    renderMenuItem "PARÁMETROS", "formAdminParams"
  end if
end function

' Custom verbs

function handleCustomVerbs
  handleCustomVerbs = true
  select case verb
    case "main": renderMain
    case "jBookingTypeOptions": jBookingTypeOptions getNumericParam("neighborId"), getNumericParam("bookingResourceId")
    case "jBookingStartOptions": jBookingStartOptions getNumericParam("neighborId"), getNumericParam("bookingResourceId"), _
      getNumericParam("bookingTypeId")
    case "jBookingDurationOptions": jBookingDurationOptions getNumericParam("neighborId"), getNumericParam("bookingResourceId"), _
      getNumericParam("bookingTypeId"), getNumericParam("bookingStart")
    case else
      if getLoggedUsrData then
        select case verb
          case "menu": renderMenu
          case "notifyNeighbors": notifyNeighbors
          case else handleCustomVerbs = false
        end select
      else
        handleCustomVerbs = false
      end if
  end select
end function

%>

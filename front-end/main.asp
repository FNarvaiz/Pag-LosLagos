<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001" %>
<% option explicit %>

<!--#include file="const.asp"-->
<!--#include file="usr.asp"-->
<!--#include file="login.asp"-->
<!--#include file="classifieds.asp"-->
<!--#include file="surveys.asp"-->
<!--#include file="suppliers.asp"-->
<!--#include file="timeLine.asp"-->
<!--#include file="downloads.asp"-->
<!--#include file="bookings.asp"-->
<!--#include file="utils/db.asp"-->
<!--#include file="forms/formsData.asp"-->
<!--#include file="forms/formsUtils.asp"-->
<%

dim serverApp: serverApp = "front-end/main.asp"

function renderMainMenu
  if getUsrData then
    %>
    <div id="espacio"></div>
    <div onclick="menu()" class="menu_bar">
      <span>Vecinos del los Lagos</span>
    </div>
    <ul id="navegador">
      <%
      dim labels
      labels = eval("mainMenuLabels" & lang)
      dim i
      for i = 0 to uBound(labels)
'        if mainMenuContents(i) <> "classifieds" or usrName = "gbd" then
          %>
          <li id="mainMenuOption<%= i %>" class="li" onclick="load(this, '<%= mainMenuContents(i) %>')"><%= labels(i) %></li>
          
          <%
      next
      %>
    </ul>
    <%
  else
    %>
    <div id="espacio"></div>
    <div onclick="menu()" class="menu_bar">
      <span>Vecinos del los Lagos</span>
    </div>
    <ul id="navegadorLogin">
      <li class="btn anchor" onclick="showLoginDialog()" id="loginMenuItem"><%= eval("loginMenuItem" & lang) %></li>
      <li class="btn anchor" onclick="load(this, 'registrationDialog')"><%= eval("registrationMenuItem" & lang) %></li>
      <li class="btn anchor" onclick="load(this, 'faq')"><%= eval("faqMenuItem" & lang) %></li>
    </ul>
    <%
  end if
end function

function renderUserMenu
  if not getUsrData then exit function
  dim logoutLabel, changePasswordLabel
  logoutLabel = eval("logoutLabel" & lang)
  changePasswordLabel = eval("changePasswordLabel" & lang)
  %>
  <span id="userName">Familia <%= usrName %></span>
  <span>|</span>
  <span class="anchor" onclick="showChangePasswordDialog()"><%= changePasswordLabel %></span>
  <span>|</span>
  <span class="anchor" onclick="logout()"><%= logoutLabel %></span>
  <%
end function

const newsCategoryId = 4

function renderHome
  if not getUsrData then exit function
  %>
  <div id="homeBg"></div>
  <div id="homePanel">
    <table cellpadding="0" cellspacing="3" width="100%" id="homeNews">
      <tr><th colspan="2" align="left" valign="middle">NOVEDADES</th></tr>
      <%
      if dbGetData("SELECT TOP 10 ID, NOMBRE, dbo.DATE_TO_STR(REC_FECHA, " & sQuotes(lang) & ") AS FECHA, CONTENIDO_TEXTO, ARCHIVO_FILENAME, ARCHIVO_FILESIZE " & _
          "FROM LINEA_TIEMPO WHERE APROBADO=1 AND ID_CATEGORIA=" & newsCategoryId & " AND CONTENIDO_TEXTO IS NOT NULL" & _
          " ORDER BY ANIO DESC, MES DESC, REC_FECHA DESC, ID DESC") then
        do while not rs.EOF
          %>
          <tr>
            <td width="80" align="left" valign="middle"><%= replace(rs("FECHA"), "/", "-") %></td>
            <td valign="middle"><span class="anchor"
              onclick="track('Novedades', '<%= rs("NOMBRE") %>', 'Familia ' + usrName); showDataViewer(<%= rs("ID") %>)"><%= rs("NOMBRE") %>
            </td>
          </tr>
          <%
          rs.MoveNext
        loop
      else
        %>
        <tr><td align="center" height="25">(No hay novedades en este momento.)</td></tr>
        <%
      end if
      dbReleaseData
      %>
    </table>
  </div>
  <div id="schedule" >
    <img src="../contenidos/iconos/android-icon-192x192.png">
 </div>
  <%
  logActivity "Inicio", ""
end function

function renderFAQ
  dim url: url = "contenidos/faq/faq.html"
  %>
  <div id="dynPanelBg"></div>
  <div id="contentsPanel">
    <div id="wrapper">
      <div id="qs">
        <h1><img src="contenidos/faq/signo.png" /> PREGUNTAS FRECUENTES</h1>
        <table width="800" cellpadding="10" cellspacing="5" >
          <tr>
            <td width="50%" class="preg" valign="top">
              <h2>Olvidé mi contraseña. ¿Cómo la recupero?</h2>
              <p>Es común no recordar este tipo de cosas.<br>Para recuperarla, en la pantalla de ingreso al sistema 
              hacé clic en “Olvidé mi contraseña”. Allí escribí el e-mail con el que te registraste, y el sistema te enviará la contraseña.</p>
            </td>
            <td class="preg" width="50%" valign="top">
              <h2>Si poseo más de un lote, ¿puedo visualizar los datos de todos ellos a la vez?</h2>
              <p>No. Cada usuario representa a un lote. Por ello, al poseer más de uno deberás ingresar de manera individual
              para visualizar la información de cada uno.</p>
            </td>
          </tr>
          
          <tr>
            <td colspan="1" class="preg" width="50%" valign="top">
              <h2>No puedo ingresar al sistema. ¿Qué hago?</h2>
              <p>Antes que nada asegúrate de ingresar correctamente tu e-mail y tu contraseña.
              Si continúas experimentando problemas, por favor, ponete en contacto con nosotros.</p>
            </td>
          </tr>
          <tr>
            <td colspan="1" class="preg" width="50%" valign="top">
              <h2>¿Cómo cambio mi contraseña?</h2>
              <p>Una vez que hayas ingresado al sistema, al pie de la pantalla hacé clic en “Cambiar contraseña”.
              Luego ingresá tu contraseña actual y la nueva que hayas elegido.
              Recordá que es aconsejable cambiar periódicamente la contraseña.</p>
            </td>
          </tr>
          
        </table>
      </div>
    </div>
  </div>
  <div id="contactPanel">
    Envianos tu consulta a <a href="mailto:info@vecinosloslagos.com.ar">info@vecinosloslagos.com.ar</a><br>
    o mediante el siguiente formulario.
    <form name="contactForm" action="<%= serverApp %>">
      <input type="hidden" name="content" value="sendContactMessage">
      <input type="hidden" name="sessionId" value="<%= request("sessionId") %>">
      <input type="hidden" name="lang" value="<%= lang %>">
      <input type="hidden" name="trackingLabel" value="mensaje de contacto">
      <table cellpadding="0" cellspacing="6" class="contact">
        <tr style="display: none">
          <td align="right">Comisión</td>
          <td>
            <select name="commissionId">
              <option selected="selected" value="">(seleccionar)</option>
              <%
              dbGetData("SELECT ID, NOMBRE FROM COMISIONES ORDER BY NOMBRE")
              do while not rs.EOF
                %>
                <option value="<%= rs("ID") %>"><%= rs("NOMBRE") %></option>
                <%
                rs.moveNext
              loop
              dbReleaseData
              %>
            </select>
          </td>
        </tr>
        <tr>
          <td align="right">Asunto</td>
          <td><input name="subject" type="text" size="40" maxlength="100"></td>
        </tr>
        <tr>
          <td align="right" valign="top">Mensaje</td>
          <td><textarea name="message" rows="6" cols="31"></textarea></td>
        </tr>
        <tr>
          <td colspan="2" align="right"><input type="button" value="Enviar" style="width: 60px"
            onclick="sendFormData('contactForm')"></td>
        </tr>
      </table>
    </form>
    </center>
  </div>
  <%
  logActivity "FAQ", ""
end function

function renderAbout
  if not getUsrData then exit function
  dim url: url = "contenidos/about/about.html"
  %>
  <div id="dynPanelBg"></div>
  <!--<iframe id="aboutContentPanel" scrolling="no" src="<= url %>"></iframe>-->
  <div class="aboutContentPanel" id="aboutResponsive">
    <section id="tabla">
      <article class="artAbout">
        <title>Comisión Arquitectura</title>
        <p>Paola Baglietto (L.90), Guadalupe Roberi (L.174)</p>
        <p><a href="mailto:arquitecturaloslagos@gmail.com" >arquitecturaloslagos@gmail.com</a></p>
      </article>
      <article class="artAbout">
        <title>Comisión Laguna</title>
        <p>Matías Bonta (L.168), Héctor Filice (L.158)</p>
        <p><a href="mailto:lagunaloslagos@gmail.com" >lagunaloslagos@gmail.com</a></p>
      </article>
      <article class="artAbout">
        <title>Comisión Seguridad</title>
        <p>Marcelo Torassa (L.286), Martin Guinart (L.234), Fernando Mauro (L.314), Alberto Monte (L.86)</p>
        <p><a href="mailto:segloslagos@gmail.com" >segloslagos@gmail.com</a></p>
      </article>
      <article class="artAbout">
        <title>Comisión Comunicación</title>
        <p>Florencia De Bella (L.323), Diego Sanchez Cavalieri (L.184), Martín Isola (L.13)</p>
        <p><a href="mailto:comunicacionloslagos@gmail.com" >comunicacionloslagos@gmail.com></a></p>
      </article>
      <article class="artAbout">
        <title>Comisión Deportes Y Recreación</title>
        <p>Carlos Coronel (L.177), Mariano Farina (L.240), Hernan Saenz (L.263), Sergio Gagliardi (L.179), Christian Espiga (L.90)</p>
         <p><a href="mailto:deportesloslagos@gmail.com" >deportesloslagos@gmail.com</a></p>
      </article>
      <article class="artAbout">
        <title>Comisión Auditoria Y Presupuesto</title>
    <p> Jose Mario De Bella (L.322), Sebastian Schvartzman (L.211), Lucas Kichic (L.126), Pedro Sebess (L.60), Javier Fernández Arancibia (L.59)</p>
        <p><a href="mailto:presupuestoloslagos@gmail.com" >presupuestoloslagos@gmail.com</a></p>
      </article>
      <article class="artAbout">
        <title>Comisión Espacios Verdes</title>
        <p>Javier Speziale (L.200), Aníbal Montes (L.309)</p>
        <p><a href="mailto:espaciosverdesloslagos@gmail.com" >espaciosverdesloslagos@gmail.com</a></p>
      </article>
      <article class="artAbout">
        <title>Comisión Tribunal de Disciplina</title>
        <p>Titulares: Marcelo Torassa (L.286), Martín Guinart (L.234), Diego Sánchez Cavalieri (L.184). Suplentes: Federico Georgiatis, Gabriela Sebess.</p>
        <p><a href="mailto:tribunaldedisciplinaloslagos@gmail.com" >tribunaldedisciplinaloslagos@gmail.com</a></p>
      </article>
      <article class="artAbout">
        <title>Comisión Club House</title>
        <p>Claudia Rueda (L.239), Ramiro López (L.301), Fernanda Sala (L.266), Sabrina Biscardi (L.143), Lorena Martinez (L.302)</p>
        <p><a href="mailto:clubhouseloslagos@gmail.com" >clubhouseloslagos@gmail.com</a></p>
      </article>
      
        
    </section>  
  
    
  <div id="infoResponsive">
  Director: Claudio Albandoz (L.54) - <a href="mailto:loslagosdirector@gmail.com" >loslagosdirector@gmail.com</a> <br />Director suplente: Marcelo Torassa (L.286) - <a href="mailto:directorsuplenteLL@gmail.com" >directorsuplenteLL@gmail.com</a><br />Intendente: Miguel Tellarini - <a href="mailto:mtellarini@avn-nordelta.com" >mtellarini@avn-nordelta.com</a>
    
  </div>  
  
</div>
  <div id="aboutContactPanel">
    <span style="font-size: 14px">Mensajes a la Administración</span><br><br>
    &nbsp;&nbsp;&nbsp;&nbsp;Por favor, completá el siguiente formulario<br>
    &nbsp;&nbsp;&nbsp;&nbsp;a la brevedad recibirás nuestra respuesta.<br><br>
    <form name="contactForm" action="<%= serverApp %>">
      <input type="hidden" name="content" value="sendContactMessage">
      <input type="hidden" name="sessionId" value="<%= request("sessionId") %>">
      <input type="hidden" name="lang" value="<%= lang %>">
      <input type="hidden" name="trackingLabel" value="mensaje de contacto">
      <table cellpadding="0" cellspacing="6" class="contact">
        <tr>
          <td align="right">Comisión</td>
          <td>
            <select name="commissionId">
              <option selected="selected" value="">(seleccionar)</option>
              <%
              dbGetData("SELECT ID, NOMBRE FROM COMISIONES ORDER BY NOMBRE")
              do while not rs.EOF
                %>
                <option value="<%= rs("ID") %>"><%= rs("NOMBRE") %></option>
                <%
                rs.moveNext
              loop
              dbReleaseData
              %>
            </select>
          </td>
        </tr>
        <tr>
          <td align="right">Asunto</td>
          <td><input class="editbox" name="subject" type="text" size="40" maxlength="100"></td>
        </tr>
        <tr>
          <td align="right" valign="top">Mensaje</td>
          <td><textarea name="message" rows="9" cols="31"></textarea></td>
        </tr>
        <tr>
          <td colspan="2" align="right"><input class="button" type="button" value="Enviar" style="width: 60px"
            onclick="sendFormData('contactForm')"></td>
        </tr>
      </table>
    </form>
    </center>
  </div>
  <%
  logActivity "Quienes somos", ""
end function

function renderHelp
  if not getUsrData then exit function
  dim url: url = "contenidos/phones/phones.html"
  %>
  <div id="dynPanelBg"></div>
  <div id="helpBg"></div>
  <div id="helpPanel">
    <iframe id="contentsPanel" src="<%= url %>"></iframe>
  </div>
  <%
  logActivity "Ayuda", ""
end function

function servicesAllowed
  servicesAllowed = true
  if usrServicesPemission then exit function
  servicesAllowed = false
  %>
  <div id="servicesDisabled">Tu acceso a este servicio se encuentra deshabilitado por la Administración.</div>
  <%
end function

function sendContactMessage(commissionId, subject, message)
  if not getUsrData then exit function
  dbGetData("SELECT UNIDAD, EMAIL FROM VECINOS WHERE ID=" & usrId)
  dim usrUnit: usrUnit = rs("UNIDAD")
  dim usrEmail: usrEmail = rs("EMAIL")
  dbReleaseData
  message = "<div style='margin:3px 1px;border:1px solid #B39DDB; background-color:#EDE7F6;color:#000;padding: 5px; font-weight: normal;' ><h2 style='margin:6px 0px;color: #000;font-weight: normal;font-size: 17pt;' >Unidad/Lote: " & usrUnit & "<br>" & _
    "Familia: " & usrName & "</h2>" & _
    "<h4 style='margin:3px 0px;color:#000; font-weight: normal;font-size: 12pt;' >Mensaje: " & replace(message, vbLf, "<br>") &"</h4></div>"
  if dbGetData("SELECT MAILS, NOMBRE FROM COMISIONES WHERE ID=" & commissionId) then
    sendMail uCase(rs("NOMBRE")) & " - Vecinos de los Lagos", rs("MAILS"), subject, message, "Familia " & usrName, usrEmail
    logActivity "Mensaje de contacto", rs("NOMBRE")
  else
    JSONAddOpFailed
    JSONAddMessage "La comisión seleccionada no existe; el mensaje no se envió."
    JSONSend
  end if
  dbReleaseData
end function

function sendMail(toName, toEmail, subject, message, replyToName, replyToEmail)
  on error resume next
  dbLog("sendMail: toName=" & toName & ", toEmail=" & toName & ", subject=" & subject)
  dbLog("sendMail: message=" & message)
  dbLog("sendMail: replyToName=" & replyToName & ", replyToEmail=" & replyToEmail)
  dim cdoMail: set cdoMail = server.CreateObject("CDO.Message")
  dim cdoConf: set cdoConf = server.createObject("CDO.Configuration") 
  dim cdoFields: set cdoFields = cdoConf.fields
  with cdoFields
    .Item ("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
    .Item ("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "localhost"
    .Item ("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
    .Item ("http://schemas.microsoft.com/cdo/configuration/sendusername") = "info@vecinosloslagos.com.ar"
    .Item ("http://schemas.microsoft.com/cdo/configuration/sendpassword") = PassMails
    .update
  end with
  with cdoMail
    set .configuration = cdoConf
    .from = "Equipo de los lagos <info@vecinosloslagos.com.ar>"
    .subject = subject
    .HTMLBody = message
  end with
  if InStr(toEmail,",")>0 then
      cdoMail.to = toEmail
  else
      cdoMail.to = toName & " <" & toEmail & ">"
  end if
  if not isNull(replyToEmail) and not isNull(replyToName) then 
    cdoMail.Cc = replyToName & " <" & replyToEmail & ">"
  end if
  cdoMail.Send
  sendMail = err.number '0 = OK
   set cdoFields = nothing
  set cdoConf = nothing
  set cdoMail = nothing
  if err.number = 0 then
    JSONAddOpOK
    JSONAddMessage "El mensaje ha sido enviado."
  else
    JSONAddOpFailed
    JSONAddMessage "Ha ocurrido un error; el mensaje no se envió."
    JSONAdd "description", left(err.description, len(err.description) - 2)
  end if
  JSONSend
end function

dbConnect
select case getStringParam("content", 30)
	case "loginDialog": renderLoginDialog
	case "faq": renderFAQ
	case "login": login
	case "logout": logout
  case "sendPasswordRemainder": sendPasswordRemainder
	case "changePasswordDialog": renderChangePasswordDialog
  case "changePassword": doChangePassword
  case "passwordRecoveryDialog": renderPasswordRecoveryDialog
  case "sendPasswordRecoveryMessage": sendPasswordRecoveryMessage getStringParam("usr", 100)
  case "registrationDialog" renderRegistrationDialog
  case "sendRegistrationMessage": sendRegistrationMessage getIntegerParam("unit"), getStringParam("name", 100), getStringParam("email", 100)

  case "mainMenu": renderMainMenu
  case "userMenu": renderUserMenu
  case "home": renderHome
  case "about": renderAbout

  case "timeLine": renderTimeLine
  case "timeLineDetail": renderTimeLineDetail getIntegerParam("timeLineRow"), getStringParam("detailKeyValues", 50)

  case "timeLineData": sendTimeLineData getIntegerParam("recordId"), getIntegerParam("forceDownload") = 1
  case "timeLineThumbnail": sendTimeLineData getIntegerParam("recordId"), false ' sendTimeLineThumbnail getIntegerParam("recordId"), getIntegerParam("height")

  case "downloads": renderDownloads 6 'getIntegerParam("categoryId")
  case "downloadListing": renderDownloadListing 6 'getIntegerParam("categoryId")
  
  case "textContent": getTextContent getIntegerParam("contentId")

  case "services": renderServices
  case "bookings": renderBookings
  case "bookingsListing": renderBookingsListing
  case "bookingsStatus": renderBookingsStatus getIntegerParam("resourceId")
  case "jBookingStartOptions": jBookingStartOptions getIntegerParam("resourceId"), getDateTimeParam("bookingDate")
  case "bookingsTurnTypes": renderBookingsTurnTypes getIntegerParam("resourceId"), getDateTimeParam("bookingDate"), getIntegerParam("turnStart")
  case "bookingsTake": bookingsTake getIntegerParam("resourceId"), getDateTimeParam("bookingDate"), getIntegerParam("turnStart"), _ 
          getIntegerParam("turnType"), getStringParam("extraNeighborUnit", 20)
  case "bookingCancel": bookingCancel getIntegerParam("bookingId")
  case "bookingsSendRequest": bookingsSendRequest getDateTimeParam("bookingDate"), getIntegerParam("placeId"), getIntegerParam("turnId")
  
  case "surveys": renderSurveys
  case "saveSurveyVote": saveSurveyVote getIntegerParam("surveyId"), getIntegerParam("voteValue"), getIntegerParam("itemNumber")
  case "saveSurveyVoteComment": saveSurveyVoteComment getIntegerParam("surveyId"), getStringParam("voteComment", 200), getIntegerParam("itemNumber")
  case "surveyItem": renderSingleSurveyItem getIntegerParam(surveyItemId)
  case "suppliers": renderSuppliers getIntegerParam("categoryId")
  case "suppliersListing": renderSuppliersListing getIntegerParam("categoryId")
  case "saveSupplierVote": saveSupplierVote getIntegerParam("supplierId"), getIntegerParam("voteValue"), getIntegerParam("itemNumber")
  case "classifieds": renderClassifieds getIntegerParam("categoryId")
  case "classifiedsListing": renderClassifiedsListing getIntegerParam("categoryId")
  case "classifiedImage": sendClassifiedImage getIntegerParam("recordId")
  case "profile": response.write("")
  case "help": renderHelp

  case "sendContactMessage": sendContactMessage getIntegerParam("commissionId"), getStringParam("subject", 200), getStringParam("message", 2000)
  case else
    JSONAddOpFailed
    JSONAddMessage "Error interno. " & getStringParam("content", 30)
    JSONSend
end select
dbDisconnect

%>


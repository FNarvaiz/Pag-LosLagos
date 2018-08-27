<%

function ZrenderloginDialog
  %>
  <div id="loginDialog">
    <table width="100%" height="100%" cellpadding="0" cellspacing="0">
      <tr>
        <td align="center"><h3>El sitio se encuentra fuera de servicio</h3></td>
      </tr>
    </table>
  </div>
  <%
end function

function renderloginDialog
  %>
  <div id="loginDialog">
    <form id="loginFormDiv" name="loginForm" onsubmit="login(this.usr.value, this.pwd.value); return false;">
        <div>
          <label id="usuario"><%= eval("loginDialogUsrFieldLabel" & lang) %></label>
          <input placeholder="Usuario" name="usr" type="text" class="loginDialogEditbox" value="">
        </div>
        <div>
            <label id="pass"><%= eval("loginDialogPwdFieldLabel" & lang) %></label>
            <input placeholder="Clave" name="pwd" type="password" class="loginDialogEditbox" value="">
        </div>
        <span class="anchor" id="errorMessage"></span> 
        <input class="buttonIngresar" type="submit" value="<%= eval("loginDialogSubmitBtnLabel" & lang) %>" class="loginDialogBtn">
        
        <span class="anchor" id="recoveryPassword" onclick="load(null, 'passwordRecoveryDialog')"><%= eval("loginDialogPasswordRecoveryBtnLabel" & lang) %></span> 
    </form>
  </div>
   <div id="logo">
    <img src="front-end/resource/logoGrande.jpg">
  </div>
  <!--<span class="redirect" onclick="load(this, 'registrationDialog')">Todavia no tienes cuenta ></span>-->
  <span class="redirect" onclick="load(this, 'passwordRecoveryDialog')">No puedo ingresar ></span>
  <span class="redirect" onclick="load(this, 'faq')">Términos y condiciones</span>
  <div id="loginDialogImg"><img src="front-end/resource/login.png" onload="document.getElementById('main').style.visibility = 'visible'; document.loginForm.usr.focus();"></div>
  <%
end function

function renderChangePasswordDialog
  %>
  <div id="changePasswordDialog">
    <form name="changePasswordForm" onsubmit="changePassword(); return false;">
    	<center>
        <table width="100%" cellpadding="6px" cellspacing="0" border="0">
          <tr>
            <td colspan="2" align="left"><%= eval("changePasswordDialogTitle" & lang) %>
              <div style="width: 100%; height: 1px; font-size: 1px; background-color: #e0e0e0; margin-top: 8px; margin-bottom: 8px"></div> 
            </td>
          </tr>
          <tr>
            <td width="100px" align="center"><%= eval("changePasswordDialogCurrentPwdFieldLabel" & lang) %></td>
          	<td><input name="pwd" type="password" class="changePasswordDialogEditbox"></td>
          </tr>
          <tr>
            <td width="100px" align="center"><%= eval("changePasswordDialogPwd1FieldLabel" & lang) %></td>
          	<td><input name="pwd1" type="password" class="changePasswordDialogEditbox"></td>
          </tr>
          <tr>
            <td width="100px" align="center"><%= eval("changePasswordDialogPwd2FieldLabel" & lang) %></td>
          	<td><input name="pwd2" type="password" class="changePasswordDialogEditbox"></td>
          </tr>
          <tr>
            <td colspan="2" align="right">
              <span style="color: #00c080">►</span><input type="button" value="<%= eval("changePasswordDialogCancelBtnLabel" & lang) %>" 
                class="changePasswordDialogBtn" onclick="hideChangePasswordDialog()">
              <span>&nbsp;&nbsp;&nbsp;</span>
              <span style="color: #00c080">►</span><input type="submit" value="<%= eval("changePasswordDialogSubmitBtnLabel" & lang) %>"
                class="changePasswordDialogBtn">
            </td>
          </tr>
        </table>
    	</center>
    </form>
  </div>
  <div id="loginDialogImg"><img src="front-end/resource/login.png" onload="document.getElementById('main').style.visibility = 'visible'; document.changePasswordForm.pwd.focus();"></div>
  <%
end function

function doChangePassword
  dim changePasswordDialogTitle, msg, btnLabel
  changePasswordDialogTitle = eval("changePasswordDialogTitle" & lang)
  btnLabel = eval("changePasswordResultBtnLabel" & lang)
  if changePassword(getStringParam("currentPwd", 50), getStringParam("newPwd", 50)) then
    msg = eval("changePasswordResultMsgOK" & lang)
  else
    msg = msg & eval("changePasswordResultMsgFailed" & lang)
  end if
  %>
  <div id="changePasswordResultDialog">
    <center>
      <table width="100%" cellpadding="6px" cellspacing="0" border="0">
        <tr>
          <td colspan="2" align="left"><%= changePasswordDialogTitle %>
            <div style="width: 100%; height: 1px; font-size: 1px; background-color: #e0e0e0; margin-top: 8px; margin-bottom: 8px"></div> 
          </td>
        </tr>
        <tr>
          <td><%= msg %></td>
        </tr>
        <tr>
          <td colspan="2" align="right">
            <input type="button" value="<%= btnLabel %>" class="changePasswordResultDialogBtn" onclick="hideChangePasswordDialog()">
          </td>
        </tr>
      </table>
    </center>
  </div>
  <div id="loginDialogImg"><img src="front-end/resource/login.png" onload="document.getElementById('main').style.visibility = 'visible';"></div>
  <%
end function

function renderRegistrationDialog
  %>
  <div id="registrationDialog">
    <form name="registrationForm" action="<%= serverApp %>">
      <input type="hidden" name="content" value="sendRegistrationMessage">
      <input type="hidden" name="lang" value="<%= lang %>">
      <input type="hidden" name="trackingLabel" value="registracion">
      <table cellpadding="0" cellspacing="6">
        <tr>
          <td colspan="2" align="left"><%= eval("registrationDialogTitle" & lang) %>
            <div style="width: 100%; height: 1px; font-size: 1px; background-color: #e0e0e0; margin-top: 8px; margin-bottom: 8px"></div> 
          </td>
        </tr>
        <tr>
          <td align="center">UNIDAD/LOTE</td>
          <td><input class="registrationDialogEditbox" name="unit" type="text" size="6" maxlength="5"></td>
        </tr>
        <tr>
          <td align="center">FAMILIA</td>
          <td><input class="registrationDialogEditbox" name="name" type="text" size="40" maxlength="100"></td>
        </tr>
        <tr>
          <td align="center" valign="top">E-MAIL</td>
          <td><input class="registrationDialogEditbox" name="email" type="text" size="40" maxlength="100"></td>
        </tr>
        <tr>
          <td colspan="2" align="right">
            <span style="color: #00c080">►</span><input type="button" value="CANCELAR"
              class="registrationDialogBtn" onclick="document.getElementById('loginMenuItem').click();">&nbsp;&nbsp;&nbsp;&nbsp;
            <span style="color: #00c080">►</span><input type="button" value="ENVIAR"
              class="registrationDialogBtn" onclick="sendFormData('registrationForm')">
          </td>
        </tr>
      </table>
    </form>
    </center>
  </div>
  <div id="loginDialogImg"><img src="front-end/resource/login.png" onload="document.getElementById('main').style.visibility = 'visible'; document.registrationForm.unit.focus();"></div>
  <%
end function

function sendRegistrationMessage(unit, name, email)
  if isNull(unit) then
    JSONAddOpFailed
    JSONAddMessage "Unidad/Lote debe ser numérico."
    JSONSend
    exit function
  end if
  if isNull(name) then
    JSONAddOpFailed
    JSONAddMessage "Falta el nombre de familia."
    JSONSend
    exit function
  end if
  if isNull(email) then
    JSONAddOpFailed
    JSONAddMessage "Falta el e-mail."
    JSONSend
    exit function
  end if
  dim message: message = ">h2>Se ha procesado la petición. En estos dias se le dará la alta.</h2><table border=1>" & _
    "<tr><td>Unidad/Lote:</td><td>" & unit & "</td></tr>" & _
    "<tr><td>Familia:</td><td>" & name & "</td></tr>" & _
    "<tr><td>e-mail:</td><td>" & email & "</td></tr>" & _
    "</table>"
  sendMail "Altas - Vecinos de los Lagos", "mtellarini@avn-nordelta.com", "Petición de alta - Unidad/lote: " & unit, message, name, email
end function

function renderPasswordRecoveryDialog
  %>
  <div id="passwordRecoveryDialog">
    <form name="passwordRecoveryForm" action="<%= serverApp %>">
      <input type="hidden" name="content" value="sendPasswordRecoveryMessage">
      <input type="hidden" name="lang" value="<%= lang %>">
      <input type="hidden" name="trackingLabel" value="recupero de clave">
      <h4>Recuperar clave</h4>
      <p><%= eval("passwordRecoveryDialogTitle" & lang) %></p>
        <div>
          <label id="usuario"><%= eval("passwordRecoveryDialogUsrFieldLabel" & lang) %></label>
          <input  autofocus placeholder="<%= eval("passwordRecoveryDialogUsrFieldLabel" & lang) %>" name="usr" type="text" class="loginDialogEditbox" value="">
        </div>
        <span class="anchor" id="errorMessage"></span> 
        <input class="buttonIngresar" type="submit" value="<%= eval("passwordRecoveryDialogSubmitBtnLabel" & lang) %>"  class="loginDialogBtn">
        <input type="button" value="CANCELAR"
                class="buttonIngresar" onclick="document.getElementById('loginMenuItem').click();">
    </form>
  </div>
  <div id="loginDialogImg"><img src="front-end/resource/login.png" onload="document.getElementById('main').style.visibility = 'visible'; "></div>
  <%
end function

function sendPasswordRecoveryMessage(usr)
  if isNull(usr) then
    JSONAddOpFailed
    JSONAddMessage "Falta el e-mail."
    JSONSend
    exit function
  end if
  if dbGetData("SELECT * FROM VECINOS WHERE EMAIL=" & sqlValue(lCase(usr))) then
    dim message: message = "Datos registrados en el sistema:<br><br><table border=1>" & _
      "<tr><td>Unidad/Lote:</td><td>" & rs("UNIDAD") & "</td></tr>" & _
      "<tr><td>Familia:</td><td>" & rs("NOMBRE") & "</td></tr>" & _
      "<tr><td>e-mail:</td><td>" & rs("EMAIL") & "</td></tr>" & _
      "<tr><td>Contraseña:</td><td>" & rs("CLAVE") & "</td></tr>" & _
      "</table>"
    sendMail "Familia " & rs("NOMBRE"), rs("EMAIL"), "Datos de acceso al sistema", message, null, null
  else
    JSONAddOpFailed
    JSONAddMessage "El e-mail ingresado no se encuentra registrado en el sistema."
    JSONSend
  end if
end function

%>

<%

function renderProfile
  if not getUsrData then exit function
  if not servicesAllowed then exit function
  dim phonePassword
  dbGetData("SELECT CLAVE_TELEFONICA FROM VECINOS WHERE ID=" & usrId)
  phonePassword = rs(0)
  dbReleaseData
  %>
  <style>
#phonePasswordBg {
  position: absolute;
  left: 30px;
  top: 20px;
	width: 390px;
	height: 120px;
  background-color: #000000;
  border-radius: 6px;
	filter: alpha(opacity=25);
  opacity:.25;
}
#phonePasswordPanel {
  position: absolute;
  left: 50px;
  top: 30px;
	width: 350px;
	height: 100px;
  color: #ffffff;
  font-size: 13px;
  overflow: hidden;
}
#phonePasswordPanel table {
  font: inherit;
}
#phonePasswordPanel input,select,textarea {
  font-family: arial,helvetica;
  font-size: 12px;
  color: #333333;
  border: solid 1px #333333;
  height: 22px;
}
  </style>
  <div id="dynPanelBg"></div>
  <div id="phonePasswordBg"></div>
  <div id="phonePasswordPanel">
    <b>Mi Contraseña Telefónica</b><br>
    <br>Será solicitada por la Guardia para autenticar tus llamadas.<br><br>
    <form name="contactForm" action="<%= serverApp %>">
      <input type="hidden" name="content" value="savePhonePassword">
      <input type="hidden" name="sessionId" value="<%= request("sessionId") %>">
      <input type="hidden" name="lang" value="<%= lang %>">
      <input type="hidden" name="trackingLabel" value="Actualización de Contraseña Telefónica">
      <table cellpadding="0" cellspacing="6" class="contact">
        <tr>
          <td align="right">Contraseña (4 dígitos)</td>
          <td><input class="editbox" name="phonePassword" type="text" maxlength="4" value="<%= phonePassword %>" style="width: 40px; padding: 0 4px;"></td>
          <td align="right"><input class="button anchor" type="button" value="Actualizar" style="width: 80px"
            onclick="sendFormData('contactForm')"></td>
        </tr>
      </table>
    </form>
  </div>
  <%
end function

function savePhonePassword(phonePassword)
  if not getUsrData then exit function
  if not servicesAllowed then exit function
  if len(phonePassword) = 4 then
    dbExecute("UPDATE VECINOS SET CLAVE_TELEFONICA = " & sqlValue(phonePassword) & " WHERE ID=" & usrId)
    JSONAddOpOK
    logActivity "Actualización de Contraseña Telefónica", "OK"
  else
    JSONAddOpFailed
    JSONAddMessage "La contraseña debe ser de 4 dígitos."
  end if
  JSONSend
end function

%>


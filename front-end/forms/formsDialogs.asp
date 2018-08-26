<%

const dialogOKBtnLabel = "Aceptar"
const dialogCancelBtnLabel = "Cancelar"

const adminLoginDialogTitle = "INGRESO AL SISTEMA"
const adminLoginDialogUsrFieldLabel = "USUARIO"
const adminLoginDialogPwdFieldLabel = "CONTRASEÑA"
const adminLoginDialogSubmitBtnLabel = "► INGRESAR"
const adminLoginDialogPwdRecoveryQuestion = "¿OLVIDÓ SU CONTRASEÑA?"

function renderAdminLoginDialog
  %>
  <form name="loginForm" onsubmit="login(); return false;">
    <div id="adminLoginDialog">
    	<center>
        <table width="100%" cellpadding="6px" cellspacing="0" border="0">
          <tr>
            <td colspan="2" align="left"><%= adminLoginDialogTitle %>
              <div style="width: 100%; height: 1px; font-size: 1px; background-color: #000000; margin-top: 8px; margin-bottom: 8px"></div> 
            </td>
          </tr>
          <tr>
            <td width="80px" align="left"><%= adminLoginDialogUsrFieldLabel %></td>
          	<td><input name="usr" type="text" size="30" class="adminLoginDialogEditbox"></td>
          </tr>
          <tr>
            <td width="80px" align="left"><%= adminLoginDialogPwdFieldLabel %></td>
          	<td><input name="pwd" type="password" size="12" class="adminLoginDialogEditbox"></td>
          </tr>
          <tr>
            <td colspan="2" align="right">
              <input type="submit" value="<%= adminLoginDialogSubmitBtnLabel %>" class="adminLoginDialogBtn">
            </td>
          </tr>
          <tr>
            <td colspan="2" align="left">
              <div style="width: 100%; height: 1px; font-size: 1px; background-color: #000000;  margin-top: 4px; margin-bottom: 8px"></div> 
              <span class="anchor" onclick="sendPasswordRemainder()"><%= adminLoginDialogPwdRecoveryQuestion %></span>
            </td>
          </tr>
        </table>
    	</center>
    </div>
  </form>
  <%
end function

const fileUploadDialogTitle = "Cargar"
const fileUploadDialogFileFieldLabel = "Archivo"

function renderFileUploadDialog
  %>
  <div id="fileUploadDialog">
    <form name="formFileUpload" encType="multipart/form-data" action="front-end/forms/utils/fileUpload.asp" method="POST" target="fileUploadFrame">
      <table class="dialog" cellpadding="0" cellspacing="0" width="100%" height="100%">
        <tr><td class="dialogTitle"><%= fileUploadDialogTitle %>&nbsp;<%= getStringParam("dataType", 30) %></td></tr>
        <tr>
          <td class="dialogBody" height="35" align="center" valign="middle">
            <%= fileUploadDialogFileFieldLabel %>&nbsp;
            <input type="file" size="70" name="uploadFormFileField">
            <input class="hidden" type="text" size="200" name="tableName" value="<%= formTable %>">
            <input class="hidden" type="text" size="200" name="keyfields" value="<%= keyFieldNames %>">
            <input class="hidden" type="text" size="200" name="keyValues" value="<%= keyFieldValues %>">
            <input class="hidden" type="text" size="200" name="recordId" value="<%= recordId %>">
            <input class="hidden" type="text" size="200" name="dbFieldBaseName" value="<%= getStringParam("dbFieldBaseName", 30) %>">
          </td>
        </tr>
        <tr>
          <td class="dialogButtons" valign="top">
            <table width="100%" cellpadding="0" cellspacing="0">
              <tr>
                <td width="45%" align="right" valign="top">
                  <input type="button" value="<%= dialogOKBtnLabel %>" onclick="fileUploadOK('<%= formId %>');">
                </td>
                <td width="10%">&nbsp;</td>
                <td width="45%" align="left" valign="top">
                  <input type="button" value="<%= dialogCancelBtnLabel %>" onclick="fileUploadCancel('<%= formId %>');">
                </td>
              </tr>
            </table>
          </td>
        </tr>
      </table>
    </form>
    <iframe name="fileUploadFrame" id="fileUploadFrame" src="" onload="fileUploaded('<%= formId %>');"></iframe>
  </div>
  <%
end function

function renderHTMLEditorDialog
  %>
  <div id="htmlEditorDialog">
    <table class="dialog" cellpadding="0" cellspacing="0" width="100%" height="100%">
      <tr><td class="dialogTitle" id="htmlEditorTitle"></td></tr>
      <tr>
        <td class="dialogBody" height="430" align="center" valign="middle">
          <div id="htmlEditorDiv" style="width: 100%; height: 430px"></div>
        </td>
      </tr>
      <tr>
        <td class="dialogButtons" valign="top">
          <table width="100%" cellpadding="0" cellspacing="0">
            <tr>
              <td width="45%" align="right" valign="top">
                <input type="button" value="<%= dialogOKBtnLabel %>" onclick="htmlEditorOK();">
              </td>
              <td width="10%">&nbsp;</td>
              <td width="45%" align="left" valign="top">
                <input type="button" value="<%= dialogCancelBtnLabel %>" onclick="htmlEditorCancel();">
              </td>
            </tr>
          </table>
        </td>
      </tr>
    </table>
  </div>
  <%
end function

%>

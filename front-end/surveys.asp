<%

function renderSurveys
  if not getUsrData then exit function
  if not servicesAllowed then exit function
  %>
  <div id="dynPanelBg"></div>
  <div id="surveysTitle">
    IDEAS Y PROPUESTAS PARA MEJORAR EL FUNCIONAMIENTO DEL BARRIO
    <p>En esta sección proponemos encuestas para valorar diversas ideas y propuestas para la mejora y mantenimiento del barrio.</p>
    <p>Estas ideas podrán presentarse en la próxima reunión para debatirlas.</p>
    <p>Para participar indicá tu valoración en cada ítem. El resultado que aparece es el promedio alcanzado por cada propuesta.</p>
  </div>
  <div id="surveysPanel">
    <div id="surveysLeftCol">
    <%
      dbGetData("SELECT ID, dbo.NOMBRE_CATEGORIA_ENCUESTA(ID_CATEGORIA) AS CATEGORIA, FECHA_ALTA, NOMBRE, " & _
        "COALESCE(DESCRIPCION, '') AS DESCRIPCION, " & _
        "dbo.VALORACION_ENCUESTA(ID) AS VALORACION, dbo.NOMBRE_VALORACION(dbo.VALORACION_ENCUESTA(ID)) AS NOMBRE_VALORACION, " & _ 
        "dbo.CANTIDAD_VALORACIONES_ENCUESTA(ID) AS CANTIDAD_VALORACIONES, " & _
        "dbo.NOMBRE_VALORACION(dbo.VALORACION_ENCUESTA_VECINO(ID, " & usrId & ")) AS VALORACION_VECINO, " & _
        "dbo.COMENTARIO_ENCUESTA_VECINO(ID, " & usrId & ") AS COMENTARIO_VALORACION " & _
        "FROM ENCUESTAS WHERE VIGENTE=1 " & _
        "ORDER BY FECHA_ALTA DESC")
      dim i: i = 1
      do while not rs.EOF
        %>
        <div class="surveyItem" id="surveysItem<%= i %>"><% renderSurveyItem(i) %></div>
        <%
        i = i + 1
        rs.moveNext
      loop
      dbReleaseData
      %>
    </div>
    <div id="surveysRightCol">
    </div>
  </div>
  <%
  logActivity "Encuestas", ""
end function

function renderSurveyItem(itemNumber)
  %>
  <div class="surveyResult">
  <%
    if isNull(rs("VALORACION")) then
      %>
      <div class="voteQty">Sin votos</div>
      <%
    else
      %>
      <div class="voteResultBg"><img src="front-end/resource/votos_bg.png"></div>
      <div class="voteResultfg">
        <img src="front-end/resource/votos_fg.png"
          style="position: absolute; clip:rect(0px,<%= cInt(round(64 * cDbl(rs("VALORACION")) / 100))  %>px,10px,0px)">
      </div>
      <div class="voteQty">
        <%= rs("NOMBRE_VALORACION") %><br>Votos: <%= rs("CANTIDAD_VALORACIONES") %>
      </div>
      <%
    end if
  %>
  </div>
  <div class="surveyTitle">
    </span><%= rs("NOMBRE") %><span class="surveyDate"><%= month(rs("FECHA_ALTA")) %>-<%= year(rs("FECHA_ALTA")) %>
  </div>
  <div class="surveyDescription"><%= replace(rs("DESCRIPCION"), vbLf, "<br>") %></div>
  <div class="surveyVote">
  <%
    dim s: s = rs("VALORACION_VECINO")
    dim t: t = ""
    dim confirmationQuestion: confirmationQuestion = ""
    if isNull(s) then 
      s = ""
    else
      confirmationQuestion = "Para " & ucase(rs("NOMBRE")) & _
        ", ¿querés cambiar tu valoración de " & ucase(s) & " a "
    end if
    if s = "Malo" then
      %>
      <div class="voteMarked"><img src="front-end/resource/voto1.png" title="Malo"></div>
      <%
    else
      if len(s) > 0 then t = confirmationQuestion & ucase("Malo") & "?"
      %>
      <div class="vote"><img src="front-end/resource/voto1.png" title="Malo" 
        onclick="saveSurveyVote('<%= t %>', <%= rs("ID") %>, 0, <%= itemNumber %>)"></div>
      <%
    end if
    if s = "Regular" then
      %>
      <div class="voteMarked"><img src="front-end/resource/voto2.png" title="Regular"></div>
      <%
    else
      if len(s) > 0 then t = confirmationQuestion & ucase("Regular") & "?"
      %>
      <div class="vote"><img src="front-end/resource/voto2.png" title="Regular" 
        onclick="saveSurveyVote('<%= t %>', <%= rs("ID") %>, 33, <%= itemNumber %>)"></div>
      <%
    end if
    if s = "Bueno" then
      %>
      <div class="voteMarked"><img src="front-end/resource/voto3.png" title="Bueno"></div>
      <%
    else
      if len(s) > 0 then t = confirmationQuestion & ucase("Bueno") & "?"
      %>
      <div class="vote"><img src="front-end/resource/voto3.png" title="Bueno"  
        onclick="saveSurveyVote('<%= t %>', <%= rs("ID") %>, 67, <%= itemNumber %>)"></div>
      <%
    end if
    if s = "Excelente" then
      %>
      <div class="voteMarked"><img src="front-end/resource/voto4.png" title="Excelente"></div>
      <%
    else
      if len(s) > 0 then t = confirmationQuestion & ucase("Excelente") & "?"
      %>
      <div class="vote"><img src="front-end/resource/voto4.png" title="Excelente"  
        onclick="saveSurveyVote('<%= t %>', <%= rs("ID") %>, 100, <%= itemNumber %>)"></div>
      <%
    end if
    if isNull(rs("VALORACION_VECINO")) then
      %>
      <div class="vote">Aún no expresaste tu valoración.</div>
      <%
    else
      %>
      <div class="vote">Tu valoración: <%= rs("VALORACION_VECINO") %>.</div>
      <div class="voteComment" id="voteComment<%= itemNumber %>" style="margin-top: 6px;">Comentario:<br>
        <textarea onfocus="this.select()" 
          style="width: 330px; height: 50px; background-color: #484848; color: #e2e2e2; padding: 0 4px;" 
          placeholder=" Podés dejarnos aquí un comentario..."
          maxlength="140" id="voteCommentText<%= itemNumber %>"><%= rs("COMENTARIO_VALORACION") %></textarea>
        <input type="button" value="Guardar" class="anchor"  
          style="vertical-align: top; background-color: #999; border: 1px solid; border-radius: 8px;"
          onclick="saveSurveyVoteComment('Confirmá si querés guardar este comentario.', <%= rs("ID") %>, <%= itemNumber %>)">
      </div>
      <%
    end if
  %>
  </div>
  <%
end function

function saveSurveyVote(surveyId, voteValue, itemNumber)
  if not getUsrData then exit function
  dim voteId: voteId = null
  if dbGetData("SELECT ID FROM ENCUESTAS_VALORACIONES WHERE ID_ENCUESTA=" & surveyId & " AND ID_VECINO=" & usrId) then
    voteId = rs(0)
  end if
  dbReleaseData
  if isNull(voteId) then
    dbExecute("INSERT INTO ENCUESTAS_VALORACIONES (ID_ENCUESTA, ID_VECINO, VALORACION) VALUES (" & surveyId & ", " & usrId & ", " & voteValue & ")")
  else
    dbExecute("UPDATE ENCUESTAS_VALORACIONES SET VALORACION=" & voteValue & ", FECHA=GETDATE() WHERE ID_ENCUESTA=" & surveyId & " AND ID_VECINO=" & usrId)
  end if
  dbGetData("SELECT ID, dbo.NOMBRE_CATEGORIA_ENCUESTA(ID_CATEGORIA) AS CATEGORIA, FECHA_ALTA, NOMBRE, " & _
    "COALESCE(DESCRIPCION, '') AS DESCRIPCION, " & _
    "dbo.VALORACION_ENCUESTA(ID) AS VALORACION, dbo.NOMBRE_VALORACION(dbo.VALORACION_ENCUESTA(ID)) AS NOMBRE_VALORACION, " & _ 
    "dbo.CANTIDAD_VALORACIONES_ENCUESTA(ID) AS CANTIDAD_VALORACIONES, " & _
    "dbo.NOMBRE_VALORACION(dbo.VALORACION_ENCUESTA_VECINO(ID, " & usrId & ")) AS VALORACION_VECINO, " & _
    "dbo.COMENTARIO_ENCUESTA_VECINO(ID, " & usrId & ") AS COMENTARIO_VALORACION " & _
    "FROM ENCUESTAS WHERE ID=" & surveyId)
  renderSurveyItem(itemNumber)
  logActivity "Voto encuesta", "Encuesta = " & rs("NOMBRE") & ", Valoración = " & rs("VALORACION_VECINO")
  dbReleaseData
end function

function saveSurveyVoteComment(surveyId, voteComment, itemNumber)
  if not getUsrData then exit function
  dim voteId: voteId = null
  if dbGetData("SELECT ID FROM ENCUESTAS_VALORACIONES WHERE ID_ENCUESTA=" & surveyId & " AND ID_VECINO=" & usrId) then
    voteId = rs(0)
  end if
  dbReleaseData
  if not isNull(voteId) then
    dbExecute("UPDATE ENCUESTAS_VALORACIONES SET COMENTARIO=" & sqlValue(voteComment) & ", FECHA=GETDATE() WHERE ID_ENCUESTA=" & surveyId & " AND ID_VECINO=" & usrId)
  end if
  dbGetData("SELECT ID, dbo.NOMBRE_CATEGORIA_ENCUESTA(ID_CATEGORIA) AS CATEGORIA, FECHA_ALTA, NOMBRE, " & _
    "COALESCE(DESCRIPCION, '') AS DESCRIPCION, " & _
    "dbo.VALORACION_ENCUESTA(ID) AS VALORACION, dbo.NOMBRE_VALORACION(dbo.VALORACION_ENCUESTA(ID)) AS NOMBRE_VALORACION, " & _ 
    "dbo.CANTIDAD_VALORACIONES_ENCUESTA(ID) AS CANTIDAD_VALORACIONES, " & _
    "dbo.NOMBRE_VALORACION(dbo.VALORACION_ENCUESTA_VECINO(ID, " & usrId & ")) AS VALORACION_VECINO, " & _
    "dbo.COMENTARIO_ENCUESTA_VECINO(ID, " & usrId & ") AS COMENTARIO_VALORACION " & _
    "FROM ENCUESTAS WHERE ID=" & surveyId)
  renderSurveyItem(itemNumber)
  logActivity "Comentario encuesta", "Encuesta = " & rs("NOMBRE")
  dbReleaseData
end function

%>

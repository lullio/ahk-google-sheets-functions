gui, add, listview, x1 y1 w1175 r10 grid vLVAll,name|year_start|year_end|position|height|height (f)|height (in)|height (m)|weight|weight (kg)|LMD (kg/m)|birth_date|college
Gui, Show

; tipoExportacao = CSV, HTML, JSON
GS_GetAllData(url:="https://docs.google.com/spreadsheets/d/1aBohx1LumhF6UICZgnI6iao8YjgK72_qnfU8O-_szqo/edit#gid=731134197", tipoExportacao:="CSV", tabNameOrId:="", rangeData := "", sqlQuery:="") {

   RegExMatch(url, "\/d\/(.+)\/", sheetId) ; capturar o id da url da planilha
   RegExMatch(url, "#gid=(.+)", sheetTabName) ; capturar sheet id ou sheet name da planilha

   If(!sheetTabName)
      ; se não passou o argumento na função, vai tentar capturar o id pela url da planilha
      sheetTabName := sheetTabName1
   Else
      ; caso tenha passado o argumento na função especificando o nome da tab ou id
      sheetTabName := tabNameOrId

   sheetUrl := % "https://docs.google.com/spreadsheets/d/" sheetId1 "/gviz/tq?tqx=out:" tipoExportacao "&range=" rangeData "&gid=" sheetTabName "&tq=" GS_EncodeDecodeURI(sqlQuery)

   ; Faz a solicitação para obter dados
   httpRequest := ComObjCreate("WinHttp.WinHttpRequest.5.1")
   httpRequest.Open("GET", sheetUrl, true)
   httpRequest.Send("") ; o corpo da solicitação precisa estar vazio
   httpRequest.WaitForResponse()

   ; Verifica a resposta
   if (httpRequest.Status == 200) {
      responseBody := httpRequest.ResponseText
      ; Exibe a resposta (pode ser necessário analisar o JSON)
      ;    MsgBox % responseBody
      msgbox % responseBody
      ; Notify("Sucesso ao deletar a propriedade!",responseBody,"","Duration=15 Image=300 GC=0xddddff BC=0x0000FF")
      Clipboard := responseBody
      return responseBody
   } else {
      erro := httpRequest.ResponseText
      msgbox % erro
      ; Notify("Error", "Ocorreu um erro ao deletar a propriedade!", 25, "Style=Info")
      ; Notify("Error", erro, 5, "Style=Info")
      ; MsgBox % httpRequest.ResponseText
      return ""
   }
}

lst := "", cnt := 0
GS_SearchRows(spreadsheetData, SearchTerm:=".*", listviewVar:= "LVAll"){
   
   GuiControl, -Redraw, % listviewVar
   for x,y in strsplit(spreadsheetData,"`n","`r")
      if (RegExMatch(y, "im)" SearchTerm) && x>1) ; x>1 para nao pegar o header (?<!https:\/\/www\.)notion
      {
         row := [], ++cnt
         loop, parse, y, CSV ; dividir a linha em células
            ; if (a_index <= 13)																	;or if a_index in 1,4,5
            row.push(a_loopfield)
         LV_add("",row*)
      }
   ; SB_SetText("Match(es) da última Pesquisa: " cnt, 4)

   GuiControl, Focus, % listviewVar ; dar foco na listview após pesquisar
   LV_Modify(1, "+Select") ; selecionar primeiro item da listview

   loop, % LV_GetCount("col")
      LV_ModifyCol(a_index,"AutoHdr")
   
   GuiControl, +Redraw, % listviewVar

   If(LV_GetCount() = 0)
      MsgBox, 4112 , Erro!, A Pesquisa não retornou nada`nAtualizando...!, 2
}
GS_SearchColumns(spreadsheetData, SearchTerm:=".*", listviewVar:= "LVAll"){

   ; Divide a primeira linha da variável spreadsheetData` (que contém os dados da planilha) em colunas usando a vírgula como delimitador. Serve para capturar apenas a primeira linha, que é o header, onde possui os nomes das colunas
   msgbox % firstCarriageReturnPos := instr(spreadsheetData, "`n")
   msgbox % firstLine := substr(spreadsheetData, 1, firstCarriageReturnPos - 1)
   msgbox % columns := strsplit(firstLine, ",")
   ; for index, column in columns
   ;    MsgBox, % "Coluna " index ": " column
   msgbox % strsplit(substr(spreadsheetData, 1, instr(spreadsheetData,"`n`")-1),",")
   ; for x,y in strsplit(substr(spreadsheetData, 1, instr(spreadsheetData,"`n")-1),",")
   for index, column in columns ; 
      ; Encontra a posição da coluna cujo nome corresponde ao valor digitado pelo usuário (needle)
      InStr(column, SearchTerm) && pos := index
      ; RegExMatch(SearchTerm, y) && pos := X
   msgbox % "position ===========" pos
   GuiControl, -Redraw, % listviewVar
   for x,y in strsplit(spreadsheetData,"`n","`r")
      loop, parse, y, CSV
         if (x>1 && a_index = pos)
            LV_add("",a_loopfield)
   LV_ModifyCol(1,"AutoHdr")
   GuiControl, +Redraw, % listviewVar
}

GS_EncodeDecodeURI(str, encode := true, component := true) {
   static Doc, JS
   if !Doc {
      Doc := ComObjCreate("htmlfile")
      Doc.write("<meta http-equiv=""X-UA-Compatible"" content=""IE=9"">")
      JS := Doc.parentWindow
      ( Doc.documentMode < 9 && JS.execScript() )
   }
   Return JS[ (encode ? "en" : "de") . "codeURI" . (component ? "Component" : "") ](str)
}

; https://docs.google.com/spreadsheets/d/1aBohx1LumhF6UICZgnI6iao8YjgK72_qnfU8O-_szqo/edit#gid=731134197
; https://docs.google.com/spreadsheets/d/1YcrxDb6w00PutmGKt-1NOcrFZyECOr5ej16OyEA8ZSQ/edit#gid=1398015473


allData := GS_GetAllData()
GS_SearchColumns(allData, "name")
; GS_SearchRows(allData)

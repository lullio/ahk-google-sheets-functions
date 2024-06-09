gui, add, listview, x1 y1 w1175 r10 grid vLVAll,name|year_start|year_end|position|height|height (f)|height (in)|height (m)|weight|weight (kg)|LMD (kg/m)|birth_date|college
; gui, add, listview, x1 y1 w1175 r10 grid vLVAll
Gui, Show

/*
   * Retornar todos os dados de uma planilha google sheet
   * tipoExportacao = CSV, HTML, JSON
*/
GS_GetAllData(url:="https://docs.google.com/spreadsheets/d/1UPX_JCNXP6wcdPadQI3mtXgZah7pSSPtxk11sjijMZE/edit?usp=sharing", tipoExportacao:="CSV", tabNameOrId:="", rangeData := "", sqlQuery:="") {

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
      MsgBox % responseBody

      for index, row in strsplit(responseBody,"`n","`r")
      {
         ; Pula a primeira linha
         if (A_Index = 1)
            continue

         ; Remove aspas extras
         row := RegExReplace(row, """")
         ; Divide a linha em campos
         fields := StrSplit(row, ",")
         ; Adiciona os campos ao ListView
         LV_Add("", fields*)
      }

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

/*
   * Pesquisa por termo e insere na ListView somente as linhas que derem match
*/
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
/*
; * procura por uma coluna específica e atualiza a listview para mostrar somente 1 coluna
; * as outras colunas ficam vazias
*/
GS_SearchColumns(spreadsheetData, SearchTerm:=".*", listviewVar:= "LVAll"){

   ; Divide a primeira linha da variável spreadsheetData` (que contém os dados da planilha) em colunas usando a vírgula como delimitador. Serve para capturar apenas a primeira linha, que é o header, onde possui os nomes das colunas
   firstCarriageReturnPos := instr(spreadsheetData, "`n") ; encontrar posição da quebra de linha
   firstLine := substr(spreadsheetData, 1, firstCarriageReturnPos - 1) ; json da primeira linha(colunas)
   columns := strsplit(firstLine, ",") ; dividir a primeira linha em colunas
   ; msgbox % strsplit(substr(spreadsheetData, 1, instr(spreadsheetData,"`n`")-1),",")
   ; for x,y in strsplit(substr(spreadsheetData, 1, instr(spreadsheetData,"`n")-1),",")
   for index, column in columns ;
      ; Encontra a posição da coluna cujo nome corresponde ao valor digitado pelo usuário (SearchTerm)
      InStr(column, SearchTerm) && pos := index
   ; RegExMatch(SearchTerm, y) && pos := X
   ; msgbox % "position ===========" pos
   GuiControl, -Redraw, % listviewVar
   for x,y in strsplit(spreadsheetData,"`n","`r")
      loop, parse, y, CSV
         if (x>1 && a_index = pos)
            LV_add("",a_loopfield)

   loop, % LV_GetCount("col")
      LV_ModifyCol(a_index,"AutoHdr")
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

      /*
         * Procurar por um termo em todas as colunas e adicionar as linhas correspondentes
      */
GS_SearchRowsByColumnValue(spreadsheetData, SearchTerm:="", listviewVar:= "LVAll") {
   ; Passo 1: Capturar a primeira linha da planilha (header)
   firstCarriageReturnPos := InStr(spreadsheetData, "`n")
   firstLine := SubStr(spreadsheetData, 1, firstCarriageReturnPos - 1)
   columns := StrSplit(firstLine, ",")

   ; Passo 2: Preparar a ListView
   GuiControl, -Redraw, %listviewVar%
   LV_Delete()

   ; Adicionar headers à ListView
   ; for index, column in columns
   ;     LV_Add("", column)

   ; Passo 3: Procurar por SearchTerm em todas as colunas e exibir linhas correspondentes
   rows := StrSplit(SubStr(spreadsheetData, firstCarriageReturnPos + 1), "`n")
   for rowIndex, row in rows {
       rowData := StrSplit(row, ",")
       found := false

       ; Verificar cada coluna na linha
       for colIndex, cellData in rowData {
           if (InStr(cellData, SearchTerm)) {
               found := true
               break
           }
       }

       ; Se encontrado, adicionar a linha à ListView
       if (found) {
           LV_Add("", rowData*)
       }
   }

   loop, % LV_GetCount("col")
      LV_ModifyCol(a_index,"AutoHdr")
   GuiControl, +Redraw, %listviewVar%
}

      /*
         * Procurar pelo nome de coluna específica(ex: year_end) e em seguida procurar por um termo específico nessa coluna (ex:1990), adicionar somente as linhas que correspondem ao valor nessa coluna
      */
GS_SearchRowsBySpecificColumnAndValue(spreadsheetData, SearchTerm:="1990", columnName:="year_end", listviewVar:= "LVAll") {
   ; Passo 1: Separar a primeira linha para termos todas as colunas(headers)
   firstCarriageReturnPos := InStr(spreadsheetData, "`n")
   firstLine := SubStr(spreadsheetData, 1, firstCarriageReturnPos - 1)
   columns := StrSplit(firstLine, ",")
   columnIndex := -1

   ; ENCONTRAR A POSIÇÃO DA COLUNA QUE É IGUAL A "X" (columnName)
   Loop, % columns.MaxIndex() {
       if (InStr(columns[A_Index], columnName)) {
           columnIndex := A_Index
           msgbox % A_Index
           break
       }
   }

   if (columnIndex = -1) {
       MsgBox, 48, Error, Column "%columnName%" not found.
       return
   }

   ; Passo 2: Preparar a ListView
   GuiControl, -Redraw, %listviewVar%
   LV_Delete()

   ; ; Adicionar header à ListView
   ; LV_Add("", columnName)

   ; Passo 3: Procurar por SearchTerm apenas na coluna específica
   rows := StrSplit(SubStr(spreadsheetData, firstCarriageReturnPos + 1), "`n")
   for rowIndex, row in rows {
       rowData := StrSplit(row, ",")

       ; Verificar se o valor na coluna específica corresponde a SearchTerm
       cellData := rowData[columnIndex]
       if (InStr(cellData, SearchTerm)) {
           LV_Add("", rowData*)
       }
   }

   ; Ajustar coluna da ListView
   LV_ModifyCol("AutoHdr")
   GuiControl, +Redraw, %listviewVar%
}

allData := GS_GetAllData()
; GS_SearchRowsBySpecificColumnAndValue(allData, "1990")
; GS_SearchRows(allData)

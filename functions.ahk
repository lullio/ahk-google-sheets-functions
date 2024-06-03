; tipoExportacao = CSV, HTML, JSON
; 
GS_GetAllData(url, tipoExportacao:="CSV") {
   
   RegExMatch(url, "\/d\/(.+)\/", sheetId)
   ; msgbox %capture_sheetURL_key1%
   RegExMatch(url, "#gid=(.+)", sheetTabName)

   sheetUrl := % "https://docs.google.com/spreadsheets/d/" sheetId1 "/gviz/tq?tqx=out:" tipoExportacao "&range=" rangeData "&gid=" sheetTabName1 "&tq=" GS_EncodeDecodeURI(PlanilhaQuery)

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
       Notify("Sucesso ao deletar a propriedade!",responseBody,"","Duration=15 Image=300 GC=0xddddff BC=0x0000FF")
       Clipboard := responseBody
       return responseBody
   } else {
        erro := httpRequest.ResponseText
        Notify("Error", "Ocorreu um erro ao deletar a propriedade!", 25, "Style=Info") 
        Notify("Error", erro, 5, "Style=Info") 
        ; MsgBox % httpRequest.ResponseText
       return ""
   }
}


; GS_SearchRows(VarPesquisarDados,PlanilhaLink, PlanilhaQuery, PlanilhaTipoExportacao, PlanilhaRange, PlanilhaNomeId){
;    ; PlanilhaLink := checkSpreadsheetLink(PlanilhaLink)
;    cnt := 0
;    Gui Submit, NoHide
;    planilha := GS_GetCSV(PlanilhaLink, PlanilhaQuery, PlanilhaTipoExportacao, PlanilhaRange, PlanilhaNomeId)
;    ; msgbox % planilha
;    GuiControl, -Redraw, LVAll
;    LV_Delete()
;    for x,y in strsplit(planilha,"`n","`r")
;       ; if instr(y,VarPesquisarDados) ; se encontrar o texto digitado no searchbox na linha
;       ; if RegExMatch(y, "im).*" VarPesquisarDados ".*") ; se encontrar o texto digitado no searchbox na linha
;       if (RegExMatch(y, "im)" VarPesquisarDados) && x>1) ; x>1 para nao pegar o header (?<!https:\/\/www\.)notion
;          {
;          row := [], ++cnt
;          loop, parse, y, CSV ; dividir a linha em células
;                ; if (a_index <= 13)																	;or if a_index in 1,4,5
;                row.push(a_loopfield)
;          LV_add("",row*)
;          }
;    SB_SetText("Match(es) da última Pesquisa: " cnt,  4)
;    ; loop, % lv_getcount("col")
;    ; LV_ModifyCol(a_index,"AutoHdr")
;    ; LV_ModifyCol(1, "30 right")
;    GuiControl, +Redraw, LVAll
;    GuiControl, Focus, LVAll ; dar foco na listview após pesquisar
;    LV_Modify(1, "+Select") ; selecionar primeiro item da listview
;    LV_ModifyCol()
;    i++
;    If(LV_GetCount() = 0){
;       MsgBox, 4112 , Erro!, A Pesquisa não retornou nada`nAtualizando...!, 2
;       GS_GetListView_Update(PlanilhaLink, PlanilhaQuery, PlanilhaTipoExportacao, PlanilhaRange, PlanilhaNomeId)
;       ; Sleep, 500
;       ; Notify().AddWindow("Erro",{Time:3000,Icon:28,Background:"0x990000",Title:"ERRO",TitleSize:15, Size:15, Color: "0xCDA089", TitleColor: "0xE1B9A4"},"w330 h30","setPosBR")
;       GuiControl, Focus, BtnPesquisar ; dar foco no botao
;    }
; }


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

msgbox % GS_GetAllData("https://docs.google.com/spreadsheets/d/1YcrxDb6w00PutmGKt-1NOcrFZyECOr5ej16OyEA8ZSQ/edit#gid=1398015473")

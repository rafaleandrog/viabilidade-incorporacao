# viabilidade-incorporacao

## Google Apps Script — doPost handler

Para integrar o salvamento de terrenos na planilha Google, adicione o seguinte trecho ao `doPost` do Apps Script, dentro do bloco `if/else if` que verifica `action`:

```javascript
} else if (action === "saveTerrain") {
  var tSheet = ss.getSheetByName("terrenos");
  if (!tSheet) {
    tSheet = ss.insertSheet("terrenos");
    tSheet.appendRow(["id","createdAt","nome","cidade","estado","projeto","etapa","areaGleba","areaApp","kmlNome"]);
  }
  tSheet.appendRow([
    payload.id || "", payload.createdAt || new Date().toISOString(),
    payload.nome || "", payload.cidade || "", payload.estado || "",
    payload.projeto || "", payload.etapa || "",
    payload.areaGleba || 0, payload.areaApp || 0, payload.kmlNome || ""
  ]);
}
```

> `fotoBase64` e `kmlBase64` ficam apenas no localStorage (tamanho proibitivo para Sheets).

function exportProjectToJSON() {
  var scriptId = ScriptApp.getScriptId();
  var url = "https://script.googleapis.com/v1/projects/" + scriptId + "/content";

  var response = UrlFetchApp.fetch(url, {
    method: "get",
    headers: {
      Authorization: "Bearer " + ScriptApp.getOAuthToken()
    }
  });

  var content = JSON.parse(response.getContentText());

  var exportData = {
    scriptId: scriptId,
    exportedAt: new Date().toISOString(),
    files: content.files || []
  };

  var json = JSON.stringify(exportData, null, 2);

  var html = `
    <html>
      <body>
        <a id="download" download="appscript-project.json"
           href="data:application/json;charset=utf-8,${encodeURIComponent(json)}">
        </a>
        <script>
          document.getElementById('download').click();
        </script>
      </body>
    </html>
  `;

  var output = HtmlService.createHtmlOutput(html);
  SpreadsheetApp.getUi().showModalDialog(output, "Download JSON");
}
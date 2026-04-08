function doGetPagina(e) {
  const pagina = (e && e.parameter && e.parameter.p) ? e.parameter.p : 'frontend_integracao';
  const template = HtmlService.createTemplateFromFile(pagina);
  template.apiUrl = ScriptApp.getService().getUrl();
  return template.evaluate()
    .setTitle('Cobertura de Escalas · SAD BH')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

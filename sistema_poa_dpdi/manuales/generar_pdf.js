// Genera los manuales en PDF a partir de manuales/fuente/*.html
//   node manuales/generar_pdf.js          (requiere Node.js y Playwright con Chromium)
const path = require('path');
const fs = require('fs');
const { chromium } = require('playwright');

const FUENTE = path.join(__dirname, 'fuente');
const SALIDA = path.join(__dirname, 'pdf');
const MANUALES = [
  ['manual_usuario.html', 'Manual_Usuario_POA_DPDI.pdf'],
  ['manual_administrador.html', 'Manual_Administrador_POA_DPDI.pdf'],
  ['manual_tecnologia.html', 'Manual_Tecnologia_POA_DPDI.pdf'],
];
const MESES = ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Setiembre', 'Octubre', 'Noviembre', 'Diciembre'];
const hoy = new Date();
const FECHA = `${MESES[hoy.getMonth()]} ${hoy.getFullYear()}`;

(async () => {
  fs.mkdirSync(SALIDA, { recursive: true });
  const navegador = await chromium.launch();
  const pagina = await navegador.newPage();
  for (const [html, pdf] of MANUALES) {
    const fuente = fs.readFileSync(path.join(FUENTE, html), 'utf8').replaceAll('{{FECHA}}', FECHA);
    const temporal = path.join(FUENTE, `.tmp_${html}`);
    fs.writeFileSync(temporal, fuente);
    try {
      await pagina.goto('file://' + temporal, { waitUntil: 'networkidle' });
      await pagina.pdf({
        path: path.join(SALIDA, pdf), format: 'A4', printBackground: true, preferCSSPageSize: true });
      console.log('PDF generado:', pdf);
    } finally {
      fs.unlinkSync(temporal);
    }
  }
  await navegador.close();
})();

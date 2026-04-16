const http = require('http');
const fs = require('fs');
const path = require('path');

const PORT = Number(process.env.PORT) || 5000;
const HOST = '0.0.0.0';

const RUTAS = {
  '/': 'index.html',
  '/formulario': 'formulario.html',
  '/formulario/': 'formulario.html'
};

const SECURITY_HEADERS = {
  'Content-Type': 'text/html; charset=utf-8',
  'X-Content-Type-Options': 'nosniff',
  'Referrer-Policy': 'strict-origin-when-cross-origin',
  'Permissions-Policy': 'camera=(), microphone=(), geolocation=()'
};

const server = http.createServer((req, res) => {
  const url = decodeURIComponent(req.url.split('?')[0]);
  const archivo = RUTAS[url];

  if (!archivo) {
    res.writeHead(404, SECURITY_HEADERS);
    res.end('Página no encontrada');
    return;
  }

  const filePath = path.join(__dirname, archivo);

  fs.readFile(filePath, (err, data) => {
    if (err) {
      res.writeHead(404, SECURITY_HEADERS);
      res.end('Página no encontrada');
      return;
    }
    res.writeHead(200, SECURITY_HEADERS);
    res.end(data);
  });
});

server.listen(PORT, HOST, () => {
  console.log(`Servidor corriendo en http://${HOST}:${PORT}/`);
  console.log(`Formulario público: http://${HOST}:${PORT}/formulario`);
});

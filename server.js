const http = require('http');
const fs = require('fs');
const path = require('path');

const PORT = 5000;
const HOST = '0.0.0.0';

const RUTAS = {
  '/':           'index.html',
  '/formulario': 'formulario.html'
};

const server = http.createServer((req, res) => {
  const url = req.url.split('?')[0];
  const archivo = RUTAS[url] || 'index.html';
  const filePath = path.join(__dirname, archivo);

  fs.readFile(filePath, (err, data) => {
    if (err) {
      res.writeHead(404);
      res.end('Página no encontrada');
      return;
    }
    res.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8' });
    res.end(data);
  });
});

server.listen(PORT, HOST, () => {
  console.log(`Servidor corriendo en http://${HOST}:${PORT}/`);
  console.log(`Formulario público: http://${HOST}:${PORT}/formulario`);
});

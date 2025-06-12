const express = require('express');
const https = require('https');
const fs = require('fs');
const path = require('path');

const app = express();

// Servir les fichiers statiques
app.use(express.static('.'));

// Headers CORS
app.use((req, res, next) => {
    res.header('Access-Control-Allow-Origin', '*');
    res.header('Access-Control-Allow-Methods', 'GET, POST, PUT, DELETE');
    res.header('Access-Control-Allow-Headers', 'Content-Type');
    next();
});

// Configuration HTTPS - Updated to use local certificates
const options = {
    key: fs.readFileSync('key.pem'),
    cert: fs.readFileSync('cert.pem')
};

https.createServer(options, app).listen(3000, () => {
    console.log('HTTPS Server running on https://localhost:3000');
});
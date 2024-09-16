// app.js
const express = require('express');
const multer = require('multer');
const xlsx = require('xlsx');
const mysql = require('mysql2');
const path = require('path');
const bodyParser = require('body-parser');

const app = express();

// Configuración de Multer para subir archivos
const storage = multer.diskStorage({
    destination: (req, file, cb) => {
        cb(null, 'uploads/');
    },
    filename: (req, file, cb) => {
        cb(null, file.originalname);
    }
});
const upload = multer({ storage: storage });

// Configuración de EJS y Body-Parser
app.set('view engine', 'ejs');
app.use(bodyParser.urlencoded({ extended: false }));
app.use(express.static(path.join(__dirname, 'public')));

// Conexión a la base de datos MySQL
const connection = mysql.createConnection({
    host: 'localhost',
    user: 'root',
    password: '',
    database: 'bd_inmobilariogpa'
});

connection.connect((err) => {
    if (err) throw err;
    console.log('Conectado a la base de datos MySQL.');
});

// Rutas
app.get('/', (req, res) => {
    res.render('index');
});

app.get('/upload', (req, res) => {
    res.render('upload');
});

app.post('/upload', upload.single('excelFile'), (req, res) => {
    const filePath = req.file.path;
    const workbook = xlsx.readFile(filePath);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const data = xlsx.utils.sheet_to_json(worksheet);

    connection.query('TRUNCATE TABLE gestion_comercial', (err) => {
        if (err) throw err;

        const query = 'INSERT INTO gestion_comercial (usuario_comercial,cedula_cliente, nombre_cliente, fecha_asignacion, numero_caso, valor_financiado, matricula, subclasificacion, regional, responsable, estado, comercial, resumen_gestion) VALUES ?';
        const values = data.map(row => [
            row.usuario_comercial,
            row.cedula_cliente,
            row.nombre_cliente,
            row.fecha_asignacion,
            row.numero_caso,
            row.valor_financiado,
            row.matricula,
            row.subclasificacion,
            row.regional,
            row.responsable,
            row.estado,
            row.comercial,
            row.resumen_gestion
        ]);

        connection.query(query, [values], (err) => {
            if (err) throw err;
            res.send('Datos importados exitosamente.');
        });
    });
});

// Iniciar el servidor
app.listen(3000, () => {
    console.log('Servidor iniciado en el puerto 3000.');
});
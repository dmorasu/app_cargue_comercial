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
app.use((err, req, res, next) => {
    console.error(err.stack);
    res.status(500).send('Ocurrió un error inesperado');
});

// Conexión a la base de datos MySQL
const connection = mysql.createConnection({
    host: '192.168.4.6',
    user: 'desarrolloti',
    password: 'd3cr3t05',
    database: 'bd_inmobiliario',
    waitForConnections: true,
    connectionLimit: 10,
    queueLimit: 0,
    connectTimeout: 10000,
    keepAliveInitialDelay: 10000
});
// Habilitar keep-alive en el pool
connection.on('connection', function (connection) {
    connection.ping();  // Envía un ping para mantener la conexión activa
  });

connection.connect((err) => {
    if (err) {
        console.error('Error al conectar a la base de datos:', err);
        process.exit(1); // Finaliza el proceso si no se puede conectar
    } else {
        console.log('Conectado a la base de datos MySQL.');
    }
});

// Pool de conexiones para la base de datos
connection.on('error', function(err) {
    console.log('Error en la base de datos:', err);
    if (err.code === 'PROTOCOL_CONNECTION_LOST') {
      // En caso de pérdida de conexión, intenta reconectar
      console.log('Conexión perdida. Intentando reconectar...');
      pool.getConnection((error, connection) => {
        if (error) {
          console.error('Error al reconectar:', error);
        } else {
          console.log('Reconexión exitosa.');
          connection.release();
        }
      });
    } else {
      throw err;  // Lanza otros errores
    }
  });


// Función para convertir fechas de Excel a formato MySQL (YYYY-MM-DD)
function excelDateToJSDate(excelDate) {
    // Verificar si la fecha es un número (formato de serie de Excel)
    if (typeof excelDate === 'number') {
        const daysSinceEpoch = excelDate - 25569; // El sistema de fechas de Excel empieza el 1 de enero de 1970
        const date = new Date(daysSinceEpoch * 86400 * 1000); // Convertir días en milisegundos
        return date.toISOString().split('T')[0]; // Devolver en formato YYYY-MM-DD
    }
    // Si la fecha ya está en formato de texto, devolverla directamente
    return excelDate;
}

// Rutas
app.get('/', (req, res) => {
    res.render('index');
});

app.get('/upload', (req, res) => {
    res.render('upload');
});

app.post('/upload', upload.single('excelFile'), (req, res) => {
    if (!req.file) {
        return res.status(400).send('No se ha subido ningún archivo.');
    }

    const filePath = req.file.path;
    let workbook;
    try {
        workbook = xlsx.readFile(filePath);
    } catch (error) {
        return res.status(400).send('Error al leer el archivo Excel.');
    }

    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const data = xlsx.utils.sheet_to_json(worksheet);
    
    if (data.length === 0) {
        return res.status(400).send('El archivo Excel está vacío.');
    }

    connection.query('TRUNCATE TABLE gestion_comercial', (err) => {
        if (err) throw err;

        const query = 'INSERT INTO gestion_comercial (usuario_comercial, cedula_cliente, nombre_cliente, fecha_asignacion, numero_caso, valor_financiado, matricula, subclasificacion, regional, responsable, estado, comercial, resumen_gestion) VALUES ?';
        const values = data.map(row => [
            row.usuario_comercial,
            row.cedula_cliente,
            row.nombre_cliente,
            excelDateToJSDate(row.fecha_asignacion), // Conviertes la fecha aquí
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
            if (err) return res.status(500).send('Error al insertar los datos en la base de datos.');
            res.send('Datos importados exitosamente.');
        });
    });
});

// Iniciar el servidor
app.listen(3000, () => {
    console.log('Servidor iniciado en el puerto 3000.');
});

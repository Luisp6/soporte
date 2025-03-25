const express = require('express');
const path = require('path');
const fs = require('fs');
const ExcelJS = require('exceljs');

const app = express();
app.use(express.json());

// Servir archivos estáticos (index.html, dashboard.html, etc.) desde "public"
app.use(express.static(path.join(__dirname, 'public')));

// Ruta del archivo JSON para persistir tickets
const TICKETS_FILE = path.join(__dirname, 'tickets.json');

/* =====================
   FUNCIONES DE LECTURA/ESCRITURA CON MANEJO DE ERRORES
===================== */
function loadTickets() {
  try {
    if (!fs.existsSync(TICKETS_FILE)) {
      // Si no existe, retornamos array vacío
      console.log('Archivo tickets.json no existe. Se creará al guardar.');
      return [];
    }
    const data = fs.readFileSync(TICKETS_FILE, 'utf-8');
    return JSON.parse(data);
  } catch (error) {
    // Si ocurre un error (permisos, JSON corrupto, etc.), retornamos array vacío
    console.log('Error leyendo tickets.json:', error.message);
    return [];
  }
}

function saveTickets(tickets) {
  try {
    fs.writeFileSync(TICKETS_FILE, JSON.stringify(tickets, null, 2));
  } catch (error) {
    // Si hay error (permisos, etc.), lo mostramos pero no detenemos el servidor
    console.log('Error escribiendo tickets.json (cambios no guardados):', error.message);
  }
}

function getNextTicketId(tickets) {
  if (tickets.length === 0) {
    return 1;
  }
  const maxId = Math.max(...tickets.map(t => t.id));
  return maxId + 1;
}

/* =====================
   ENDPOINTS
===================== */

// Crear ticket o insumo
app.post('/api/tickets', (req, res) => {
  try {
    const tickets = loadTickets();
    const { owner, role, type, asunto, descripcion, attachment } = req.body;
    // Fecha en formato ISO para filtrar por mes (YYYY-MM)
    const fecha = new Date().toISOString();
    const ticketId = getNextTicketId(tickets);

    const newTicket = {
      id: ticketId,
      owner,
      role,
      type, // "incidencia" o "insumo"
      asunto,
      descripcion,
      fecha,
      estado: 'Abierto',
      adminResponse: '',
      attachment // base64 o null
    };

    tickets.push(newTicket);
    saveTickets(tickets);

    res.status(201).json({
      message: 'Ticket/pedido creado con éxito.',
      ticket: newTicket
    });
  } catch (error) {
    console.error('Error al crear ticket:', error);
    res.status(500).json({ error: 'Error al crear ticket.' });
  }
});

// Obtener lista de tickets
app.get('/api/tickets', (req, res) => {
  const tickets = loadTickets();
  res.json(tickets);
});

// Actualizar ticket (estado/respuesta admin)
app.patch('/api/tickets/:id', (req, res) => {
  try {
    const ticketId = parseInt(req.params.id, 10);
    const { nuevoEstado, adminResponse } = req.body;
    const tickets = loadTickets();
    const ticket = tickets.find(t => t.id === ticketId);
    if (!ticket) {
      return res.status(404).json({ error: 'Ticket no encontrado.' });
    }

    if (nuevoEstado !== undefined) {
      ticket.estado = nuevoEstado;
    }
    if (adminResponse !== undefined) {
      ticket.adminResponse = adminResponse;
    }

    saveTickets(tickets);
    res.json({ message: 'Ticket actualizado.', ticket });
  } catch (error) {
    console.error('Error al actualizar ticket:', error);
    res.status(500).json({ error: 'Error al actualizar ticket.' });
  }
});

// Informe en Excel de incidencias por mes (YYYY-MM)
app.get('/api/reports/incidencias', async (req, res) => {
  try {
    const mes = req.query.mes; // ejemplo: "2023-07"
    if (!mes) {
      return res.status(400).json({ error: 'El parámetro "mes" es requerido en formato YYYY-MM' });
    }
    const allTickets = loadTickets();
    // Filtrar incidencias que tengan "fecha" comenzando con mes
    const incidencias = allTickets.filter(ticket => 
      ticket.type === 'incidencia' && ticket.fecha.startsWith(mes)
    );

    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Incidencias');

    worksheet.columns = [
      { header: 'ID', key: 'id', width: 10 },
      { header: 'Usuario', key: 'owner', width: 20 },
      { header: 'Rol', key: 'role', width: 15 },
      { header: 'Asunto', key: 'asunto', width: 30 },
      { header: 'Descripción', key: 'descripcion', width: 40 },
      { header: 'Fecha', key: 'fecha', width: 25 },
      { header: 'Estado', key: 'estado', width: 15 },
      { header: 'Respuesta Admin', key: 'adminResponse', width: 30 }
    ];

    incidencias.forEach(ticket => {
      worksheet.addRow(ticket);
    });

    // Forzar descarga del Excel
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename=Incidencias_${mes}.xlsx`);
    await workbook.xlsx.write(res);
    res.end();
  } catch (error) {
    console.error('Error generando informe:', error);
    res.status(500).json({ error: 'Error generando informe.' });
  }
});

// Iniciar servidor
const PORT = 3000;
app.listen(PORT, () => {
  console.log(`Servidor iniciado en http://localhost:${PORT}`);
});

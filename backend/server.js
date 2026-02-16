const express = require('express');
const { google } = require('googleapis');
const cors = require('cors');
require('dotenv').config();

const app = express();
app.use(cors());
app.use(express.json());

const auth = new google.auth.GoogleAuth({
    keyFile: 'credentials.json', 
    scopes: ['https://www.googleapis.com/auth/spreadsheets'],
});

const SPREADSHEET_ID = '1Q_OF8Rqvni--awnOj2xRMDMVnP3789Lc-cl8MSG7VpE'; 


app.get('/api/datos/:mes', async (req, res) => {
    const { mes } = req.params;
    const sheets = google.sheets({ version: 'v4', auth });

    try {
        const response = await sheets.spreadsheets.values.get({
            spreadsheetId: SPREADSHEET_ID,
            range: `${mes}!A2:K`, 
        });
        res.json(response.data.values || []);
    } catch (error) {
        console.error("Error en Google Sheets:", error);
        res.status(500).json({ error: "No se pudo leer el Excel", detalle: error.message });
    }
});

app.post('/api/nuevo', async (req, res) => {
    const { mes, fila } = req.body;
    const sheets = google.sheets({ version: 'v4', auth });

    try {
        await sheets.spreadsheets.values.append({
            spreadsheetId: SPREADSHEET_ID,
            range: `${mes}!A2`,
            valueInputOption: 'USER_ENTERED',
            resource: { values: [fila] },
        });
        res.status(200).send("Registro guardado con éxito");
    } catch (error) {
        res.status(500).send("Error al guardar en el Excel");
    }
});

app.put('/api/editar', async(req,res)=>{
    const { mes, filaIndex, nuevosValores } = req.body;
    const sheets = google.sheets({ version: 'v4', auth });

    try{
        const rangoDestino = `${mes}!A${filaIndex + 2}`;
        
        await sheets.spreadsheets.values.update({
            spreadsheetId: SPREADSHEET_ID,
            range: rangoDestino,
            valueInputOption: 'USER_ENTERED',
            resource: { values: [nuevosValores] },
        });
        res.status(200).send("Registro actualizado correctamente");
    }
    catch (error){
        res.status(500).send("Error al editar: " + error.message);
    }
});

app.delete('/api/borrar', async(req, res)=>{
    const { mes, filaIndex } = req.body;
    const sheets = google.sheets({ version: 'v4', auth });

    try{
        const spreadsheet = await sheets.spreadsheets.get({ spreadsheetId: SPREADSHEET_ID });
        const sheetId = spreadsheet.data.sheets.find(s=> s.properties.title===mes).properties.sheetId;

        await sheets.spreadsheets.batchUpdate({
            spreadsheetId: SPREADSHEET_ID,
            resource: {
                requests: [{
                    deleteDimension: {
                        range: {
                            sheetId: sheetId,
                            dimension: "ROWS",
                            starrIndex: filaIndex + 1,
                            endIndex: filaIndex + 2
                        }
                    }
                    
                }]
            }
        });
        res.status(200).send("Fila eliminada y orden ajustado");
    } 
    catch (error){
        res.status(500).send("Error al borrar: "+ error.message);
    }
});

const PORT = 5000;
app.listen(PORT, () => {
    console.log(`SERVIDOR LISTO EN: http://localhost:${PORT}`);
});
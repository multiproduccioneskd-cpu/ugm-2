const axios = require('axios');

export default async function handler(req, res) {
    // 1. Configuración de cabeceras CORS para la tele de la U
    res.setHeader('Access-Control-Allow-Credentials', true);
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.setHeader('Access-Control-Allow-Methods', 'GET,OPTIONS');
    res.setHeader('Access-Control-Allow-Headers', 'X-CSRF-Token, X-Requested-With, Accept, Accept-Version, Content-Length, Content-MD5, Content-Type, Date, X-Api-Version');

    if (req.method === 'OPTIONS') {
        res.status(200).end();
        return;
    }

    try {
        // 2. Pedir el Token de acceso a Microsoft de forma nativa por HTTP
        const tokenUrl = `https://login.microsoftonline.com/${process.env.TENANT_ID}/oauth2/v2.0/token`;
        
        const params = new URLSearchParams();
        params.append('client_id', process.env.CLIENT_ID);
        params.append('scope', 'https://graph.microsoft.com/.default');
        params.append('client_secret', process.env.CLIENT_SECRET);
        params.append('grant_type', 'client_credentials');

        const tokenResponse = await axios.post(tokenUrl, params, {
            headers: { 'Content-Type': 'application/x-www-form-urlencoded' }
        });

        const accessToken = tokenResponse.data.access_token;

        // 3. Pegarle directo a la API de SharePoint para traer los eventos
        const graphUrl = `https://graph.microsoft.com/v1.0/sites/${process.env.SHAREPOINT_SITE_ID}/lists/${process.env.SHAREPOINT_LIST_ID}/items?expand=fields`;
        
        const graphResponse = await axios.get(graphUrl, {
            headers: { Authorization: `Bearer ${accessToken}` }
        });

        const items = graphResponse.data.value || [];

        // 4. Mapear los datos limpios para tu index.html
        const eventosProcesados = items.map(item => {
            return {
                title: item.fields.Title || item.fields.title || "Evento sin título",
                sala: item.fields.Sala || item.fields.sala || "Por definir",
                casillaTiempo: item.fields.Fecha || item.fields.casillaTiempo || item.fields.EventDate || null,
                Destinatario: item.fields.Destinatario || item.fields.destinatario || null
            };
        });

        // Forzar a Vercel a no guardar caché vieja
        res.setHeader('Cache-Control', 'no-shadow, no-store, must-revalidate');
        res.status(200).json(eventosProcesados);

    } catch (error) {
        console.error("Error en cartelera API:", error.response ? error.response.data : error.message);
        res.status(500).json({ error: "Error de conexión con SharePoint", detalle: error.message });
    }
}

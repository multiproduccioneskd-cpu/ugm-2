module.exports = async (req, res) => {
    // 1. Cabeceras CORS obligatorias para la tele de la U
    res.setHeader('Access-Control-Allow-Credentials', true);
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.setHeader('Access-Control-Allow-Methods', 'GET,OPTIONS');
    res.setHeader('Access-Control-Allow-Headers', 'X-CSRF-Token, X-Requested-With, Accept, Accept-Version, Content-Length, Content-MD5, Content-Type, Date, X-Api-Version');

    if (req.method === 'OPTIONS') {
        res.status(200).end();
        return;
    }

    try {
        // 2. Pedir el Token a Microsoft usando fetch nativo (sin librerías externas)
        const tokenUrl = `https://login.microsoftonline.com/${process.env.TENANT_ID}/oauth2/v2.0/token`;
        
        const bodyParams = new URLSearchParams();
        bodyParams.append('client_id', process.env.CLIENT_ID);
        bodyParams.append('scope', 'https://graph.microsoft.com/.default');
        bodyParams.append('client_secret', process.env.CLIENT_SECRET);
        bodyParams.append('grant_type', 'client_credentials');

        const tokenResponse = await fetch(tokenUrl, {
            method: 'POST',
            headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
            body: bodyParams.toString()
        });

        const tokenData = await tokenResponse.json();
        const accessToken = tokenData.access_token;

        if (!accessToken) {
            throw new Error("No se pudo obtener el token de acceso de Microsoft");
        }

        // 3. Consultar la lista de SharePoint de forma directa
        const graphUrl = `https://graph.microsoft.com/v1.0/sites/${process.env.SHAREPOINT_SITE_ID}/lists/${process.env.SHAREPOINT_LIST_ID}/items?expand=fields`;
        
        const graphResponse = await fetch(graphUrl, {
            method: 'GET',
            headers: { 'Authorization': `Bearer ${accessToken}` }
        });

        const graphData = await graphResponse.json();
        const items = graphData.value || [];

        // 4. Mapear los datos reales de SharePoint hacia tu index.html
        const eventosProcesados = items.map(item => {
            return {
                title: item.fields.Title || item.fields.title || "Evento sin título",
                sala: item.fields.Sala || item.fields.sala || "Por definir",
                casillaTiempo: item.fields.Fecha || item.fields.casillaTiempo || item.fields.EventDate || null,
                
                // Rescatamos la columna tipo Elección de SharePoint de forma segura
                Destinatario: item.fields.Destinatario || item.fields.destinatario || null
            };
        });

        // Desactivar caché para actualizaciones en tiempo real en la tele
        res.setHeader('Cache-Control', 'no-shadow, no-store, must-revalidate');
        res.status(200).json(eventosProcesados);

    } catch (error) {
        console.error("Error en la API de la cartelera:", error.message);
        res.status(500).json({ error: "Error de conexión con SharePoint", detalle: error.message });
    }
};

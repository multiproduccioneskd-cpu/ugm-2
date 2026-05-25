module.exports = async (req, res) => {
    // Cabeceras CORS obligatorias
    res.setHeader('Access-Control-Allow-Credentials', true);
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.setHeader('Access-Control-Allow-Methods', 'GET,OPTIONS');
    res.setHeader('Access-Control-Allow-Headers', 'X-CSRF-Token, X-Requested-With, Accept, Accept-Version, Content-Length, Content-MD5, Content-Type, Date, X-Api-Version');

    if (req.method === 'OPTIONS') {
        res.status(200).end();
        return;
    }

    try {
        // 1. Obtener Token de Acceso desde Azure Entra ID
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
            return res.status(500).json({ error: "No se pudo autenticar con Azure" });
        }

        // 2. Consulta directa expandiendo todos los campos de SharePoint de un viaje
        const graphUrl = `https://graph.microsoft.com/v1.0/sites/${process.env.SHAREPOINT_SITE_ID}/lists/${process.env.SHAREPOINT_LIST_ID}/items?expand=fields`;
        
        const graphResponse = await fetch(graphUrl, {
            method: 'GET',
            headers: { 'Authorization': `Bearer ${accessToken}` }
        });

        const graphData = await graphResponse.json();
        const items = graphData.value || [];

        // 3. Mapeo limpio tolerando mayúsculas, minúsculas y objetos tipo Choice
        const eventosProcesados = items.map(item => {
            const f = item.fields || {};
            
            let rawDestinatario = f.Destinatario || f.destinatario || null;
            if (rawDestinatario && typeof rawDestinatario === 'object') {
                rawDestinatario = rawDestinatario.Value || rawDestinatario.value || null;
            }

            return {
                title: f.Title || f.title || "Evento sin título",
                sala: f.Sala || f.sala || "Por definir",
                casillaTiempo: f.Fecha || f.fecha || f.casillaTiempo || f.EventDate || null,
                Destinatario: rawDestinatario
            };
        });

        // Respuesta limpia sin caché
        res.setHeader('Cache-Control', 'no-shadow, no-store, must-revalidate');
        res.status(200).json(eventosProcesados);

    } catch (error) {
        console.error("Error en API Cartelera:", error.message);
        res.status(500).json({ error: "Error de comunicación", detalle: error.message });
    }
};

module.exports = async (req, res) => {
    // Cabeceras CORS de siempre
    res.setHeader('Access-Control-Allow-Credentials', true);
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.setHeader('Access-Control-Allow-Methods', 'GET,OPTIONS');
    res.setHeader('Access-Control-Allow-Headers', 'X-CSRF-Token, X-Requested-With, Accept, Accept-Version, Content-Length, Content-MD5, Content-Type, Date, X-Api-Version');

    if (req.method === 'OPTIONS') {
        res.status(200).end();
        return;
    }

    try {
        // 1. Token con Azure
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
            return res.status(500).json({ error: "Error de token" });
        }

        // 2. Consulta limpia a SharePoint pidiendo el objeto fields completo en bruto
        const graphUrl = `https://graph.microsoft.com/v1.0/sites/${process.env.SHAREPOINT_SITE_ID}/lists/${process.env.SHAREPOINT_LIST_ID}/items?expand=fields`;
        
        const graphResponse = await fetch(graphUrl, {
            method: 'GET',
            headers: { 'Authorization': `Bearer ${accessToken}` }
        });

        const graphData = await graphResponse.json();
        const items = graphData.value || [];

        // 3. Mapeo simplificado: Mandamos el objeto "fields" completo en bruto hacia el HTML
        // Así el HTML puede buscar la columna como quiera sin romper el servidor
        const eventosProcesados = items.map(item => {
            return {
                title: item.fields ? (item.fields.Title || item.fields.title || "Evento sin título") : "Evento sin título",
                sala: item.fields ? (item.fields.Sala || item.fields.sala || "Por definir") : "Por definir",
                casillaTiempo: item.fields ? (item.fields.Fecha || item.fields.fecha || item.fields.EventDate || null) : null,
                
                // 🚀 AQUÍ ESTÁ EL TRUCO: Mandamos todo el contenedor fields para inspeccionarlo en el navegador
                rawFields: item.fields || {}
            };
        });

        res.setHeader('Cache-Control', 'no-shadow, no-store, must-revalidate');
        res.status(200).json(eventosProcesados);

    } catch (error) {
        console.error(error);
        res.status(500).json({ error: "Error de conexión" });
    }
};

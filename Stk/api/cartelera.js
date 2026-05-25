module.exports = async (req, res) => {
    // Cabeceras CORS limpias para la tele
    res.setHeader('Access-Control-Allow-Credentials', true);
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.setHeader('Access-Control-Allow-Methods', 'GET,OPTIONS');
    res.setHeader('Access-Control-Allow-Headers', 'X-CSRF-Token, X-Requested-With, Accept, Accept-Version, Content-Length, Content-MD5, Content-Type, Date, X-Api-Version');

    if (req.method === 'OPTIONS') {
        res.status(200).end();
        return;
    }

    try {
        // 1. Autenticación limpia con Azure para sacar el Token
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
            return res.status(500).json({ error: "Error de autenticación con Azure" });
        }

        // 2. 🚀 URL EN BRUTO CORREGIDA: Traemos absolutamente todos los "fields" sin filtros que rompan la respuesta
        const graphUrl = `https://graph.microsoft.com/v1.0/sites/${process.env.SHAREPOINT_SITE_ID}/lists/${process.env.SHAREPOINT_LIST_ID}/items?expand=fields`;
        
        const graphResponse = await fetch(graphUrl, {
            method: 'GET',
            headers: { 'Authorization': `Bearer ${accessToken}` }
        });

        const graphData = await graphResponse.json();
        const items = graphData.value || [];

        // 3. Mapeo ultra-tolerante para que el JSON arme las llaves fijas que busca tu index.html
        const eventosProcesados = items.map(item => {
            const f = item.fields || {};
            return {
                title: f.Title || f.title || "Evento sin título",
                sala: f.Sala || f.sala || "Por definir",
                fecha: f.Fecha || f.fecha || f.casillaTiempo || f.EventDate || ""
            };
        });

        res.setHeader('Cache-Control', 'no-shadow, no-store, must-revalidate');
        res.status(200).json(eventosProcesados);

    } catch (error) {
        console.error("Error en el mapeo de la API:", error);
        res.status(500).json({ error: "Error interno apuntando a SharePoint", detalle: error.message });
    }
};

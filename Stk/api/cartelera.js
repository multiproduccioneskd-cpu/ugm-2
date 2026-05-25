module.exports = async (req, res) => {
    // 1. Cabeceras CORS limpias
    res.setHeader('Access-Control-Allow-Credentials', true);
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.setHeader('Access-Control-Allow-Methods', 'GET,OPTIONS');
    res.setHeader('Access-Control-Allow-Headers', 'X-CSRF-Token, X-Requested-With, Accept, Accept-Version, Content-Length, Content-MD5, Content-Type, Date, X-Api-Version');

    if (req.method === 'OPTIONS') {
        res.status(200).end();
        return;
    }

    try {
        // 2. Rescatar Token de Microsoft por HTTP Puro usando variables de entorno
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
            return res.status(500).json({ error: "No se obtuvo token de Azure" });
        }

        // 3. Pegarle a la lista de SharePoint pidiendo expandir todos los campos ('fields')
        const graphUrl = `https://graph.microsoft.com/v1.0/sites/${process.env.SHAREPOINT_SITE_ID}/lists/${process.env.SHAREPOINT_LIST_ID}/items?expand=fields($select=Title,title,Sala,sala,Fecha,casillaTiempo,EventDate,Destinatario,destinatario)`;
        
        const graphResponse = await fetch(graphUrl, {
            method: 'GET',
            headers: { 'Authorization': `Bearer ${accessToken}` }
        });

        const graphData = await graphResponse.json();
        
        // Si no viene la propiedad value, dejamos un arreglo vacío de respaldo
        const items = graphData.value || [];

        // 4. Mapear de forma ultra-flexible para tolerar cualquier minúscula/mayúscula de SharePoint
        const eventosProcesados = items.map(item => {
            const f = item.fields || {};
            
            return {
                title: f.Title || f.title || "Evento sin título",
                sala: f.Sala || f.sala || "Por definir",
                casillaTiempo: f.Fecha || f.casillaTiempo || f.EventDate || null,
                Destinatario: f.Destinatario || f.destinatario || null
            };
        });

        // 5. Entregar la data real y limpia directo al index.html
        res.setHeader('Cache-Control', 'no-shadow, no-store, must-revalidate');
        res.status(200).json(eventosProcesados);

    } catch (error) {
        console.error("Error en la ejecución de la API:", error.message);
        res.status(500).json({ error: "Error al conectar con SharePoint", detalle: error.message });
    }
};

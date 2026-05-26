module.exports = async (req, res) => {
    // Cabeceras CORS esenciales para la tele de la U
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.setHeader('Access-Control-Allow-Headers', 'Content-Type, Authorization');
    res.setHeader('Access-Control-Allow-Methods', 'GET, OPTIONS');

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

        const tokenRes = await fetch(tokenUrl, {
            method: 'POST',
            headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
            body: bodyParams.toString()
        });

        if (!tokenRes.ok) throw new Error("Error obteniendo Token de Azure");
        const tokenData = await tokenRes.json();
        const accessToken = tokenData.access_token;

        // 2. Conectar a la lista oficial expandiendo todos los fields ocultos
        const graphUrl = `https://graph.microsoft.com/v1.0/sites/ugmchile.sharepoint.com:/sites/Calen:/lists/lista%20calen/items?expand=fields&$top=100`;

        const graphRes = await fetch(graphUrl, {
            method: 'GET',
            headers: { 
                'Authorization': `Bearer ${accessToken}`,
                'Accept': 'application/json'
            }
        });

        if (!graphRes.ok) throw new Error("Microsoft Graph no respondió");
        const graphData = await graphRes.json();
        const rawItems = graphData.value || [];

        // 3. Mapeo plano directo de las columnas nativas de SharePoint
        const eventosProcesados = rawItems.map(item => {
            const f = item.fields || {};
            
            // Forzar rescate de ubicación real (U_G_M_Sala)
            let salaReal = f.U_G_M_Sala || f.Sala || f.Ubicacion || "Por definir";
            
            // Forzar rescate de fecha de calendario nativa (EventDate) o alternativas de sistema
            let fechaOriginal = f.EventDate || f.StartDate || f.EventDateTime || "";

            return {
                title: f.Title || f.LinkTitle || "Evento sin título",
                sala: salaReal,
                casillaTiempo: fechaOriginal, // Mandamos el string ISO puro de Microsoft para procesarlo en el HTML
                Destinatario: f.Destinatario || f.destinatario || ""
            };
        });

        res.setHeader('Cache-Control', 'no-shadow, no-store, must-revalidate');
        return res.status(200).json(eventosProcesados);

    } catch (error) {
        console.error("Fallo crítico backend:", error);
        return res.status(500).json({ error: error.message });
    }
};

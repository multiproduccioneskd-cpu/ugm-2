module.exports = async (req, res) => {
    // Cabeceras CORS obligatorias
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

        // 2. URL de la lista calen expandiendo fields
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

        // 3. Mapeo ultra-blindado contra campos vacíos o mañosos de SharePoint
        const eventosProcesados = rawItems.map(item => {
            const f = item.fields || {};
            
            // Buscar la SALA barriendo todas las posibilidades de SharePoint
            let salaReal = f.U_G_M_Sala || f.Sala || f.Ubicacion || f.Location || f.OData__Location || f.OData__U_G_M_Sala || "Por definir";
            
            // Buscar la FECHA barriendo todos los nombres ocultos que genera un Calendario de SharePoint
            let fechaCruda = f.EventDate || f.StartDateTime || f.StartDate || f.Fecha || f.OData__EventDate || f.OData__StartDate || "";

            // Asegurar que fechaCruda sea SIEMPRE un string válido para que el .split() del HTML no explote
            if (!fechaCruda || typeof fechaCruda !== 'string') {
                fechaCruda = "2026-05-26T00:00:00Z"; // Respaldo seguro con formato ISO por si viene null
            }

            return {
                title: f.Title || f.LinkTitle || "Evento sin título",
                sala: salaReal,
                casillaTiempo: fechaCruda, // Envía el ISO limpio que tu HTML necesita romper con .split('T')
                Destinatario: f.Destinatario || f.destinatario || ""
            };
        });

        // Desactivar caché para que los cambios en SharePoint se reflejen en tiempo real en la tele
        res.setHeader('Cache-Control', 'no-shadow, no-store, must-revalidate');
        return res.status(200).json(eventosProcesados);

    } catch (error) {
        console.error("Fallo crítico en el procesador backend:", error);
        return res.status(500).json({ error: error.message });
    }
};

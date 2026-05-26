module.exports = async (req, res) => {
    // Cabeceras CORS obligatorias para el visor de la tele
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

        // 2. 🚀 URL BLINDADA: Llamamos al endpoint plano de ítems y forzamos la carga limpia de los fields
        const graphUrl = `https://graph.microsoft.com/v1.0/sites/ugmchile.sharepoint.com:/sites/Calen:/lists/lista%20calen/items?expand=fields`;

        const graphRes = await fetch(graphUrl, {
            method: 'GET',
            headers: { 
                'Authorization': `Bearer ${accessToken}`,
                'Accept': 'application/json;odata.metadata=none' // Evita que Microsoft meta basura de sistema en el JSON
            }
        });

        if (!graphRes.ok) throw new Error("Microsoft Graph no respondió correctamente");
        const graphData = await graphRes.json();
        const rawItems = graphData.value || [];

        // 3. Mapeo directo a las variables nativas que lee tu HTML histórico
        const eventosProcesados = rawItems.map(item => {
            const f = item.fields || {};
            
            // Forzar el rescate de la sala desde el campo U_G_M_Sala de tu SharePoint
            let salaReal = f.U_G_M_Sala || f.Sala || f.Ubicacion || "Por definir";
            
            // Forzar el rescate del string de fecha/hora original (ej: "2026-05-26T16:00:00Z")
            let fechaOriginal = f.EventDate || f.StartDate || f.Fecha || "";

            return {
                title: f.Title || f.LinkTitle || "Evento sin título",
                sala: salaReal,
                
                // 🚀 CLAVE: Tu HTML histórico procesa esta variable exacta para calcular las horas locales y los filtros
                casillaTiempo: fechaOriginal 
            };
        });

        // Desactivar caché para que los cambios en SharePoint se vean de inmediato en la tele
        res.setHeader('Cache-Control', 'no-shadow, no-store, must-revalidate');
        return res.status(200).json(eventosProcesados);

    } catch (error) {
        console.error("Fallo crítico en el procesador backend:", error);
        return res.status(500).json({ error: error.message });
    }
};

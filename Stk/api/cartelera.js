// Backend Definitivo - Cartelera UGM (Sintaxis ESM Nativa para Vercel)
export default async function handler(req, res) {
    // Configuración estricta de CORS para evitar bloqueos en la tele
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
    res.setHeader('Access-Control-Allow-Methods', 'GET, OPTIONS');

    if (req.method === 'OPTIONS') {
        return res.status(200).end();
    }

    try {
        const { TENANT_ID, CLIENT_ID, CLIENT_SECRET, SITE_ID, LIST_ID } = process.env;

        // 1. Rescate del Token de Acceso desde Azure Identity
        const tokenUrl = `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`;
        const params = new URLSearchParams({
            client_id: CLIENT_ID,
            scope: 'https://graph.microsoft.com/.default',
            client_secret: CLIENT_SECRET,
            grant_type: 'client_credentials'
        });

        const tokenRes = await fetch(tokenUrl, {
            method: 'POST',
            body: params,
            headers: { 'Content-Type': 'application/x-www-form-urlencoded' }
        });
        
        if (!tokenRes.ok) {
            const errText = await tokenRes.text();
            throw new Error(`Error Azure Token: ${errText}`);
        }
        const tokenData = await tokenRes.json();
        const accessToken = tokenData.access_token;

        // 2. Consulta a Microsoft Graph por la lista de SharePoint
        const graphUrl = `https://graph.microsoft.com/v1.0/sites/${SITE_ID}/lists/${LIST_ID}/items?expand=fields`;
        const graphRes = await fetch(graphUrl, {
            headers: { 'Authorization': `Bearer ${accessToken}` }
        });

        if (!graphRes.ok) {
            const errText = await graphRes.text();
            throw new Error(`Error Microsoft Graph: ${errText}`);
        }
        const graphData = await graphRes.json();
        const rawItems = graphData.value || [];

        // 3. Mapeo Flexible e Inteligente de Columnas (Evita que el JSON caiga vacío)
        const eventosProcesados = rawItems.map(item => {
            const f = item.fields || {};
            
            // Busca la columna de la Sala bajo distintas variaciones de SharePoint
            const salaReal = f["U_x002e_G_x002e_M_x0020_Sala"] || f["Sala"] || f["Location"] || f["U_G_M_Sala"] || "Por definir";
            
            // Busca la columna de Fecha
            const fechaReal = f["EventDate"] || f["EventDateTime"] || f["StartDate"] || f["Fecha"] || "";

            return {
                title: f.Title || f.Title0 || f.LinkTitle || "Evento sin título",
                sala: salaReal,
                fecha: fechaReal
            };
        });

        // 4. Entrega de datos limpia al frontend
        return res.status(200).json(eventosProcesados);

    } catch (error) {
        console.error("Fallo crítico en Backend Cartelera:", error);
        return res.status(500).json({ 
            error: "Error interno del servidor backend", 
            mensaje: error.message 
        });
    }
}

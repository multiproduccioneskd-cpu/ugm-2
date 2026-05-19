// Backend Definitivo Reforzado - Cartelera UGM
export default async function handler(req, res) {
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
    res.setHeader('Access-Control-Allow-Methods', 'GET, OPTIONS');

    if (req.method === 'OPTIONS') return res.status(200).end();

    try {
        const { TENANT_ID, CLIENT_ID, CLIENT_SECRET, LIST_ID } = process.env;

        // 1. Obtener Token de Azure
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

        // 2. Ruta Limpia Inteligente (Busca directo por el dominio de la UGM para evitar el itemNotFound)
        const graphUrl = `https://graph.microsoft.com/v1.0/sites/ugmchile.sharepoint.com:/sites/cartelera:/lists/${LIST_ID}/items?expand=fields`;
        
        // NOTA: Si la lista no está en un sub-sitio llamado "cartelera", se usa la raíz:
        // const graphUrl = `https://graph.microsoft.com/v1.0/sites/ugmchile.sharepoint.com/lists/${LIST_ID}/items?expand=fields`;

        const graphRes = await fetch(graphUrl, {
            headers: { 'Authorization': `Bearer ${accessToken}` }
        });

        if (!graphRes.ok) {
            const errText = await graphRes.text();
            throw new Error(`Microsoft Graph respondió: ${errText}`);
        }
        
        const graphData = await graphRes.json();
        const rawItems = graphData.value || [];

        // 3. Mapeo de columnas de SharePoint
        const eventosProcesados = rawItems.map(item => {
            const f = item.fields || {};
            const salaReal = f["U_x002e_G_x002e_M_x0020_Sala"] || f["Sala"] || f["Location"] || "Por definir";
            const fechaReal = f["EventDate"] || f["EventDateTime"] || f["StartDate"] || "";

            return {
                title: f.Title || f.Title0 || "Evento sin título",
                sala: salaReal,
                fecha: fechaReal
            };
        });

        return res.status(200).json(eventosProcesados);

    } catch (error) {
        console.error("Fallo crítico:", error);
        return res.status(500).json({ 
            error: "Error interno del servidor backend", 
            mensaje: error.message 
        });
    }
}

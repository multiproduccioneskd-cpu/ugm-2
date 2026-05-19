export default async function handler(req, res) {
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
    res.setHeader('Access-Control-Allow-Methods', 'GET, OPTIONS');

    if (req.method === 'OPTIONS') return res.status(200).end();

    try {
        const { TENANT_ID, CLIENT_ID, CLIENT_SECRET } = process.env;

        // 1. Token Azure
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

        // 2. Consulta directa a la lista calen de la UGM
        const graphUrl = `https://graph.microsoft.com/v1.0/sites/ugmchile.sharepoint.com:/sites/Calen:/lists/lista%20calen/items?expand=fields&$top=100`;

        const graphRes = await fetch(graphUrl, {
            method: 'GET',
            headers: { 
                'Authorization': `Bearer ${accessToken}`,
                'Accept': 'application/json'
            }
        });

        if (!graphRes.ok) {
            const errText = await graphRes.text();
            throw new Error(`Microsoft Graph respondió: ${errText}`);
        }
        
        const graphData = await graphRes.json();
        const rawItems = graphData.value || [];

        // 3. Mapeo y Formateo Chileno Estricto
        const eventosProcesados = rawItems.map(item => {
            const f = item.fields || {};
            
            // Rastreador agresivo de columnas para pillar la Sala de la Mistral
            const salaReal = f["U_x002e_G_x002e_M_x0020_Sala"] || f["Sala"] || f["U_G_M_Sala"] || f["Location"] || f["U_x0020_G_x0020_M_x0020_Sala"] || f["Ubicacion"] || f["Ubicaci_x00f3_n"] || "Por definir";
            
            // Captura de fecha cruda de SharePoint
            const fechaCruda = f["EventDate"] || f["EventDateTime"] || f["StartDate"] || f["Fecha"] || f["Modified"] || "";
            
            let fechaFormateada = "";
            if (fechaCruda) {
                try {
                    // Forzamos que la fecha se lea en la zona horaria de Chile Continental
                    const dateObj = new Date(fechaCruda);
                    fechaFormateada = dateObj.toLocaleString("sv-SE", { timeZone: "America/Santiago" }).replace(" ", "T");
                } catch (e) {
                    fechaFormateada = fechaCruda;
                }
            }

            return {
                title: f.Title || f.Title0 || f.LinkTitle || "Evento sin título",
                sala: salaReal,
                fecha: fechaFormateada
            };
        });

        return res.status(200).json(eventosProcesados);

    } catch (error) {
        console.error("Fallo crítico backend:", error);
        return res.status(500).json({ 
            error: "Error interno del servidor backend", 
            mensaje: error.message 
        });
    }
}

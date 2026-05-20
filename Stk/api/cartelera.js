module.exports = async function handler(req, res) {
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
    res.setHeader('Access-Control-Allow-Methods', 'GET, OPTIONS');

    if (req.method === 'OPTIONS') return res.status(200).end();

    try {
        const { TENANT_ID, CLIENT_ID, CLIENT_SECRET } = process.env;

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
        
        if (!tokenRes.ok) throw new Error("Error en Token Azure");
        const tokenData = await tokenRes.json();
        const accessToken = tokenData.access_token;

        const graphUrl = `https://graph.microsoft.com/v1.0/sites/ugmchile.sharepoint.com:/sites/Calen:/lists/lista%20calen/items?expand=fields&$top=100`;
        const graphRes = await fetch(graphUrl, {
            method: 'GET',
            headers: { 'Authorization': `Bearer ${accessToken}`, 'Accept': 'application/json' }
        });

        if (!graphRes.ok) throw new Error("Microsoft Graph no respondió");
        const graphData = await graphRes.json();
        const rawItems = graphData.value || [];

        const eventosNormalizados = rawItems.map(item => {
            const f = item.fields || {};
            
            let salaReal = f.Sala || f.Ubicacion || f.Location || f.U_G_M_Sala || "Por definir";
            if (salaReal === "Por definir") {
                for (let key in f) {
                    if (typeof f[key] === 'string' && (f[key].includes("Auditorio") || f[key].includes("Gimnasio") || f[key].includes("Sala") || f[key].includes("SALÓN"))) {
                        salaReal = f[key];
                        break;
                    }
                }
            }

            // Capturamos cualquier variante de campo de fecha que use esta lista de SharePoint
            let casillaTiempo = f.EventDate || f.StartDate || f.EventDateTime || f.Fecha || f.OData__StartDate || item.createdDateTime || "";

            return {
                title: f.Title || f.LinkTitle || "Evento sin título",
                sala: salaReal,
                casillaTiempo: casillaTiempo
            };
        });

        return res.status(200).json(eventosNormalizados);

    } catch (error) {
        return res.status(500).json({ error: true, mensaje: error.message });
    }
};

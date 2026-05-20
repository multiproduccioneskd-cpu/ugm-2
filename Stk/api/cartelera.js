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

        // Consultamos la lista de SharePoint
        const graphUrl = `https://graph.microsoft.com/v1.0/sites/ugmchile.sharepoint.com:/sites/Calen:/lists/lista%20calen/items?expand=fields&$top=5`;
        const graphRes = await fetch(graphUrl, {
            method: 'GET',
            headers: { 'Authorization': `Bearer ${accessToken}`, 'Accept': 'application/json' }
        });

        if (!graphRes.ok) throw new Error("Microsoft Graph no respondió");
        const graphData = await graphRes.json();
        const rawItems = graphData.value || [];

        // HACK DE INSPECCIÓN: Mandamos las entrañas completas del primer ítem para pillar los campos ocultos
        if (rawItems.length > 0) {
            return res.status(200).json({
                _INSPECCION_RAIZ: Object.keys(rawItems[0]),
                _INSPECCION_FIELDS: rawItems[0].fields || {}
            });
        }

        return res.status(200).json({ mensaje: "Lista vacía" });

    } catch (error) {
        return res.status(500).json({ error: true, mensaje: error.message });
    }
};

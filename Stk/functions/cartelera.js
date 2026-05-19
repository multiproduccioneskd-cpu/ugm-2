exports.handler = async function(event, context) {
    if (event.httpMethod === "OPTIONS") {
        return {
            statusCode: 200,
            headers: {
                "Access-Control-Allow-Origin": "*",
                "Access-Control-Allow-Headers": "Content-Type",
                "Access-Control-Allow-Methods": "GET, OPTIONS"
            },
            body: ""
        };
    }

    try {
        const TENANT_ID = process.env.TENANT_ID;
        const CLIENT_ID = process.env.CLIENT_ID;
        const CLIENT_SECRET = process.env.CLIENT_SECRET;
        const SITE_ID = process.env.SITE_ID;
        const LIST_ID = process.env.LIST_ID;

        // 1. Obtener Token con fetch nativo
        const tokenUrl = `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`;
        const params = new URLSearchParams();
        params.append('client_id', CLIENT_ID);
        params.append('scope', 'https://graph.microsoft.com/.default');
        params.append('client_secret', CLIENT_SECRET);
        params.append('grant_type', 'client_credentials');

        const tokenRes = await fetch(tokenUrl, {
            method: 'POST',
            body: params,
            headers: { 'Content-Type': 'application/x-www-form-urlencoded' }
        });
        
        if (!tokenRes.ok) throw new Error('Error obteniendo token de Azure');
        const tokenData = await tokenRes.json();
        const accessToken = tokenData.access_token;

        // 2. Consultar SharePoint expandiendo los fields
        const graphUrl = `https://graph.microsoft.com/v1.0/sites/${SITE_ID}/lists/${LIST_ID}/items?expand=fields`;
        const graphRes = await fetch(graphUrl, {
            headers: { Authorization: `Bearer ${accessToken}` }
        });

        if (!graphRes.ok) throw new Error('Error leyendo Microsoft Graph');
        const graphData = await graphRes.json();
        const rawItems = graphData.value || [];

        // 3. Procesar mapeando las columnas de la UGM
        const eventosProcesados = rawItems.map(item => {
            const f = item.fields || {};
            let salaReal = f["U_x002e_G_x002e_M_x0020_Sala"] || f.Location || "Por definir";
            let fechaReal = f.EventDate || f.EventDateTime || f.StartDate || f.Fecha || "";

            return {
                title: f.Title || f.Title0 || f.LinkTitle || "Evento sin título",
                sala: salaReal,
                fecha: fechaReal
            };
        });

        return {
            statusCode: 200,
            headers: {
                "Access-Control-Allow-Origin": "*",
                "Access-Control-Allow-Headers": "Content-Type",
                "Access-Control-Allow-Methods": "GET, OPTIONS"
            },
            body: JSON.stringify(eventosProcesados)
        };

    } catch (error) {
        console.error("Error en la función cartelera:", error);
        return {
            statusCode: 500,
            headers: { "Access-Control-Allow-Origin": "*" },
            body: JSON.stringify({ error: error.message })
        };
    }
};

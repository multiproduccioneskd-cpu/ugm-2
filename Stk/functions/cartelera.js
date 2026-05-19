exports.handler = async function(event, context) {
    const TENANT_ID = "05970e72-c674-4f1f-8033-6e35dd7f76aa";
    const CLIENT_ID = "ceee9a3e-aa63-419c-960a-321e8726fd65";
    const CLIENT_SECRET = "Lfk8Q~-8lvREUP6Amzkd_7mdAT4Z1o16OdF8PazH";
    const SITE_ID = "ugmchile.sharepoint.com,0aaa32cc-4ad1-4d99-9ee6-78ede7cfff66,a582523d-6b22-417e-bf2b-5b8b7b7fd3e7";
    const LIST_ID = "aa1e422f-858d-4b84-a386-69ad7d09ec34";

    try {
        // 1. Obtener Token automáticamente
        const tokenUrl = `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`;
        const params = new URLSearchParams();
        params.append('client_id', CLIENT_ID);
        params.append('client_secret', CLIENT_SECRET);
        params.append('scope', 'https://graph.microsoft.com/.default');
        params.append('grant_type', 'client_credentials');

        const tokenRes = await fetch(tokenUrl, { method: 'POST', body: params, headers: { 'Content-Type': 'application/x-www-form-urlencoded' } });
        if (!tokenRes.ok) throw new Error("Error obteniendo token");
        const tokenData = await tokenRes.json();
        const token = tokenData.access_token;

        // 2. Consultar SharePoint
        const graphUrl = `https://graph.microsoft.com/v1.0/sites/${SITE_ID}/lists/${LIST_ID}/items?expand=fields`;
        const graphRes = await fetch(graphUrl, { headers: { 'Authorization': `Bearer ${token}` } });
        if (!graphRes.ok) throw new Error("Error leyendo SharePoint");
        
        const graphData = await graphRes.json();
        const rawItems = graphData.value || [];

        // 3. Procesar mapeando la columna real de la UGM
        const eventosProcesados = rawItems.map(item => {
            const f = item.fields || {};
            
            // Pescamos la columna exacta que nos reveló tu token
            let salaReal = f["U_x002e_G_x002e_M_x0020_Sala"] || f.Location || "Por definir";

            return {
                fields: {
                    Title: f.Title || "Actividad sin asunto",
                    Location: salaReal,
                    EventDate: f.EventDate || ""
                }
            };
        });

        return {
            statusCode: 200,
            headers: { 
                "Content-Type": "application/json", 
                "Access-Control-Allow-Origin": "*",
                "Cache-Control": "no-cache, no-store, must-revalidate"
            },
            body: JSON.stringify(eventosProcesados)
        };

    } catch (error) {
        return {
            statusCode: 500,
            headers: { "Content-Type": "application/json", "Access-Control-Allow-Origin": "*" },
            body: JSON.stringify({ error: error.message })
        };
    }
};
// Usamos el fetch nativo de Node.js para evitar problemas de dependencias en Vercel
export default async function handler(req, res) {
    // Configuración obligatoria de CORS para que tu GitHub Pages pueda leer los datos
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
    res.setHeader('Access-Control-Allow-Methods', 'GET, OPTIONS');

    if (req.method === 'OPTIONS') {
        return res.status(200).end();
    }

    try {
        const TENANT_ID = process.env.TENANT_ID;
        const CLIENT_ID = process.env.CLIENT_ID;
        const CLIENT_SECRET = process.env.CLIENT_SECRET;
        const SITE_ID = process.env.SITE_ID; // ¡En Vercel sí puedes usar SITE_ID!
        const LIST_ID = process.env.LIST_ID;

        // 1. Obtener Token de Acceso desde Azure Microsoft
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
        
        if (!tokenRes.ok) throw new Error('Error al solicitar token a Azure');
        const tokenData = await tokenRes.json();
        const accessToken = tokenData.access_token;

        // 2. Consultar los ítems del calendario de SharePoint
        const graphUrl = `https://graph.microsoft.com/v1.0/sites/${SITE_ID}/lists/${LIST_ID}/items?expand=fields`;
        const graphRes = await fetch(graphUrl, {
            headers: { Authorization: `Bearer ${accessToken}` }
        });

        if (!graphRes.ok) throw new Error('Error al conectar con Microsoft Graph');
        const graphData = await graphRes.json();
        const rawItems = graphData.value || [];

        // 3. Mapear y ordenar los campos críticos de la UGM
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

        // 4. Retornar el JSON limpio
        return res.status(200).json(eventosProcesados);

    } catch (error) {
        console.error("Error en API Cartelera:", error);
        return res.status(500).json({ error: "Error de comunicación backend", detalle: error.message });
    }
}

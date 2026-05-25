module.exports = async (req, res) => {
    // Cabeceras CORS esenciales para la pantalla de la tele
    res.setHeader('Access-Control-Allow-Credentials', true);
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.setHeader('Access-Control-Allow-Methods', 'GET,OPTIONS');
    res.setHeader('Access-Control-Allow-Headers', 'X-CSRF-Token, X-Requested-With, Accept, Accept-Version, Content-Length, Content-MD5, Content-Type, Date, X-Api-Version');

    if (req.method === 'OPTIONS') {
        res.status(200).end();
        return;
    }

    try {
        // 1. Obtener Token de Acceso desde Azure Entra ID
        const tokenUrl = `https://login.microsoftonline.com/${process.env.TENANT_ID}/oauth2/v2.0/token`;
        
        const bodyParams = new URLSearchParams();
        bodyParams.append('client_id', process.env.CLIENT_ID);
        bodyParams.append('scope', 'https://ugmchile.sharepoint.com/.default');
        bodyParams.append('client_secret', process.env.CLIENT_SECRET);
        bodyParams.append('grant_type', 'client_credentials');

        const tokenResponse = await fetch(tokenUrl, {
            method: 'POST',
            headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
            body: bodyParams.toString()
        });

        const tokenData = await tokenResponse.json();
        const accessToken = tokenData.access_token;

        if (!accessToken) {
            return res.status(500).json({ error: "No se pudo obtener el token de Azure" });
        }

        // 2. 🚀 URL de tu captura conectando directo a SharePoint
        const sharepointUrl = "https://ugmchile.sharepoint.com/sites/Cartelera/_api/web/lists/getbytitle('Eventos')/items";
        
        const spResponse = await fetch(sharepointUrl, {
            method: 'GET',
            headers: { 
                'Authorization': `Bearer ${accessToken}`,
                'Accept': 'application/json;odata=verbose' // Clave para que devuelva la estructura clásica d.results
            }
        });

        const spData = await spResponse.json();
        
        // Validamos la ruta nativa de datos de SharePoint
        const items = (spData.d && spData.d.results) ? spData.d.results : (spData.value || []);

        // 3. Mapeo blindado: Busca las columnas en minúsculas, mayúsculas y formato OData
        const eventosProcesados = items.map(item => {
            const f = item.fields || {};
            return {
                title: item.Title || item.title || f.Title || f.title || "Evento sin título",
                sala: item.Sala || item.OData__Sala || item.sala || f.Sala || f.sala || "Por definir",
                fecha: item.Fecha || item.OData__Fecha || item.fecha || item.EventDate || f.Fecha || ""
            };
        });

        res.setHeader('Cache-Control', 'no-shadow, no-store, must-revalidate');
        res.status(200).json(eventosProcesados);

    } catch (error) {
        console.error("Error en la API de SharePoint:", error);
        res.status(500).json({ error: "Error de conexión con la lista de SharePoint", detalle: error.message });
    }
};

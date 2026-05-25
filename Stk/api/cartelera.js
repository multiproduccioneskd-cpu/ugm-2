module.exports = async (req, res) => {
    // Cabeceras CORS obligatorias para la tele
    res.setHeader('Access-Control-Allow-Credentials', true);
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.setHeader('Access-Control-Allow-Methods', 'GET,OPTIONS');
    res.setHeader('Access-Control-Allow-Headers', 'X-CSRF-Token, X-Requested-With, Accept, Accept-Version, Content-Length, Content-MD5, Content-Type, Date, X-Api-Version');

    if (req.method === 'OPTIONS') {
        res.status(200).end();
        return;
    }

    try {
        // 1. Obtener Token de Acceso apuntando al recurso nativo de SharePoint
        const tokenUrl = `https://login.microsoftonline.com/${process.env.TENANT_ID}/oauth2/v2.0/token`;
        
        const bodyParams = new URLSearchParams();
        bodyParams.append('client_id', process.env.CLIENT_ID);
        // Usamos el scope heredado clásico de SharePoint para que dé los permisos correctos
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
            return res.status(500).json({ error: "No se pudo obtener el token de acceso" });
        }

        // 2. 🚀 El link real de tu captura: Pegarle directo a la API interna de SharePoint
        const sharepointUrl = "https://ugmchile.sharepoint.com/sites/Cartelera/_api/web/lists/getbytitle('Eventos')/items";
        
        const spResponse = await fetch(sharepointUrl, {
            method: 'GET',
            headers: { 
                'Authorization': `Bearer ${accessToken}`,
                'Accept': 'application/json;odata=verbose' // Obligatorio para que SharePoint suelte el JSON estructurado
            }
        });

        const spData = await spResponse.json();
        
        // La API nativa de SharePoint guarda los ítems en d.results
        const items = (spData.d && spData.d.results) ? spData.d.results : [];

        // 3. Mapeo limpio devolviendo las propiedades en minúscula que necesita tu index.html
        // Atajamos los nombres con OData_ por si SharePoint los renombró internamente
        const eventosProcesados = items.map(item => {
            return {
                title: item.Title || item.title || "Evento sin título",
                sala: item.Sala || item.sala || item.OData__Sala || "Por definir",
                fecha: item.Fecha || item.fecha || item.OData__Fecha || item.EventDate || ""
            };
        });

        res.setHeader('Cache-Control', 'no-shadow, no-store, must-revalidate');
        res.status(200).json(eventosProcesados);

    } catch (error) {
        console.error("Error en la API de SharePoint:", error);
        res.status(500).json({ error: "Error de conexión con la lista de SharePoint", detalle: error.message });
    }
};

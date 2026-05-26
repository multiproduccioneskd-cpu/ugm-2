module.exports = async (req, res) => {
    // Cabeceras CORS esenciales para las pantallas de la U
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.setHeader('Access-Control-Allow-Headers', 'Content-Type, Authorization');
    res.setHeader('Access-Control-Allow-Methods', 'GET, OPTIONS');

    if (req.method === 'OPTIONS') {
        res.status(200).end();
        return;
    }

    try {
        // 1. Obtener Token de Acceso desde Azure Entra ID
        const tokenUrl = `https://login.microsoftonline.com/${process.env.TENANT_ID}/oauth2/v2.0/token`;
        const bodyParams = new URLSearchParams();
        bodyParams.append('client_id', process.env.CLIENT_ID);
        bodyParams.append('scope', 'https://graph.microsoft.com/.default');
        bodyParams.append('client_secret', process.env.CLIENT_SECRET);
        bodyParams.append('grant_type', 'client_credentials');

        const tokenRes = await fetch(tokenUrl, {
            method: 'POST',
            headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
            body: bodyParams.toString()
        });

        if (!tokenRes.ok) throw new Error("Error obteniendo Token de Azure");
        const tokenData = await tokenRes.json();
        const accessToken = tokenData.access_token;

        // 2. Traer los ítems desde SharePoint
        const graphUrl = `https://graph.microsoft.com/v1.0/sites/ugmchile.sharepoint.com:/sites/Calen:/lists/lista%20calen/items?expand=fields&$top=100`;
        const graphRes = await fetch(graphUrl, {
            method: 'GET',
            headers: { 'Authorization': `Bearer ${accessToken}`, 'Accept': 'application/json' }
        });

        if (!graphRes.ok) throw new Error("Microsoft Graph no respondió");
        const graphData = await graphRes.json();
        const rawItems = graphData.value || [];

        // Obtener la fecha de hoy real en Santiago de Chile (YYYY-MM-DD)
        const hoyChileString = new Date().toLocaleDateString('sv-SE', { timeZone: 'America/Santiago' });

        // 3. 🚀 AQUÍ ESTÁ LA FÓRMULA DEL JSON SINCRONIZADO:
        const eventosProcesados = rawItems.map(item => {
            const f = item.fields || {};
            
            let salaReal = f.U_G_M_Sala || f.Sala || f.Ubicacion || "Por definir";
            let fechaCruda = f.EventDate || f.StartDateTime || f.StartDate || "";

            let fechaTexto = "9999-12-31";
            let horaTexto = "00:00";
            
            if (fechaCruda) {
                const d = new Date(fechaCruda);
                if (!isNaN(d.getTime())) {
                    // Forzamos al objeto Date a transformarse al horario nacional de Chile
                    fechaTexto = d.toLocaleDateString('sv-SE', { timeZone: 'America/Santiago' });
                    
                    // Extraemos la hora exacta en formato 24H ajustada a Santiago
                    horaTexto = d.toLocaleTimeString('es-CL', { 
                        hour: '2-digit', 
                        minute: '2-digit', 
                        hour12: false, 
                        timeZone: 'America/Santiago' 
                    });
                }
            }

            // Si por alguna razón SharePoint devuelve las 00:00 o 02:00 de un evento "Todo el día"
            if (horaTexto === "00:00" && fechaCruda.includes("T00:00")) {
                horaTexto = "Todo el día";
            }

            const esHoy = (fechaTexto === hoyChileString);

            return {
                title: f.Title || f.LinkTitle || "Evento sin título",
                sala: salaReal,
                hora: horaTexto,      // Ya viene con la hora de Chile lista
                fechaStr: fechaTexto,  // Ya viene con la fecha de Chile lista
                esHoy: esHoy,
                casillaTiempo: fechaCruda // De respaldo
            };
        });

        // Ordenar los eventos para que salgan los más temprano primero
        eventosProcesados.sort((a, b) => `${a.fechaStr}T${a.hora}`.localeCompare(`${b.fechaStr}T${b.hora}`));

        res.setHeader('Cache-Control', 'no-shadow, no-store, must-revalidate');
        return res.status(200).json(eventosProcesados);

    } catch (error) {
        console.error("Error en API:", error);
        return res.status(500).json({ error: error.message });
    }
};

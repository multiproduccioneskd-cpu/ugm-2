module.exports = async (req, res) => {
    // Cabeceras CORS limpias para la tele de la U
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.setHeader('Access-Control-Allow-Headers', 'Content-Type, Authorization');
    res.setHeader('Access-Control-Allow-Methods', 'GET, OPTIONS');

    if (req.method === 'OPTIONS') {
        res.status(200).end();
        return;
    }

    try {
        // 1. Token de Azure Entra ID
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

        // 2. URL que sí conecta a la lista calen
        const graphUrl = `https://graph.microsoft.com/v1.0/sites/ugmchile.sharepoint.com:/sites/Calen:/lists/lista%20calen/items?expand=fields`;

        const graphRes = await fetch(graphUrl, {
            method: 'GET',
            headers: { 
                'Authorization': `Bearer ${accessToken}`,
                'Accept': 'application/json'
            }
        });

        if (!graphRes.ok) throw new Error("Microsoft Graph no respondió");
        const graphData = await graphRes.json();
        const rawItems = graphData.value || [];

        // Obtener fecha de hoy real en Santiago (YYYY-MM-DD)
        const hoyChile = new Date().toLocaleDateString('sv-SE', { timeZone: 'America/Santiago' });

        // 3. Procesamos las fechas y horas tal como te gustaba antes
        const eventosProcesados = rawItems.map(item => {
            const f = item.fields || {};
            
            let salaReal = f.U_G_M_Sala || f.Sala || f.Ubicacion || "Por definir";
            let fechaCruda = f.EventDate || f.StartDate || f.EventDateTime || "";
            
            let fechaTexto = "9999-12-31";
            let horaTexto = "00:00";
            let esHoy = false;

            if (fechaCruda && typeof fechaCruda === 'string') {
                try {
                    const dateObjeto = new Date(fechaCruda);
                    if (!isNaN(dateObjeto.getTime())) {
                        // Extrae la fecha limpia ajustada a Chile
                        fechaTexto = dateObjeto.toLocaleDateString('sv-SE', { timeZone: 'America/Santiago' });
                        esHoy = (fechaTexto === hoyChile);
                        
                        // Formatea la hora en 24 horas exacta de Santiago
                        horaTexto = dateObjeto.toLocaleTimeString('es-CL', { 
                            hour: '2-digit', 
                            minute: '2-digit', 
                            hour12: false, 
                            timeZone: 'America/Santiago' 
                        });
                    }
                } catch (e) {
                    const partes = fechaCruda.split('T');
                    if (partes[0]) {
                        fechaTexto = partes[0];
                        esHoy = (fechaTexto === hoyChile);
                    }
                }
            }

            // Retornamos TODAS las variables para no romper nada en el HTML
            return {
                title: f.Title || f.LinkTitle || "Evento sin título",
                sala: salaReal,
                hora: horaTexto,
                fechaStr: fechaTexto,
                esHoy: esHoy,
                casillaTiempo: fechaCruda // De respaldo por si acaso
            };
        });

        // Ordenar cronológicamente por fecha y hora
        eventosProcesados.sort((a, b) => `${a.fechaStr}T${a.hora}`.localeCompare(`${b.fechaStr}T${b.hora}`));

        res.setHeader('Cache-Control', 'no-shadow, no-store, must-revalidate');
        return res.status(200).json(eventosProcesados);

    } catch (error) {
        console.error("Fallo crítico:", error);
        return res.status(500).json({ error: error.message });
    }
};

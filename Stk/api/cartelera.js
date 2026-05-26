module.exports = async (req, res) => {
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.setHeader('Access-Control-Allow-Headers', 'Content-Type, Authorization');
    res.setHeader('Access-Control-Allow-Methods', 'GET, OPTIONS');

    if (req.method === 'OPTIONS') {
        res.status(200).end();
        return;
    }

    try {
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

        const hoyChile = new Date().toLocaleDateString('sv-SE', { timeZone: 'America/Santiago' });

        const eventosProcesados = rawItems.map(item => {
            const f = item.fields || {};
            
            // 🔍 LOG PARA REVISAR EN VERCEL: Borra esto si quieres después, sirve para ver las columnas reales
            console.log("Campos del item:", f);

            // Intentar agarrar la sala de todas las formas posibles que genera SharePoint
            let salaReal = f.U_G_M_Sala || f.Sala || f.Ubicacion || f.Location || f.OData__Location || "Por definir";
            
            // 🔥 El gran problema de SharePoint: Las fechas cambian de nombre según el tipo de lista.
            // Buscamos en orden de probabilidad:
            let fechaCruda = f.EventDate || f.StartDateTime || f.StartDate || f.Fecha || f.OData__StartDate || "";
            
            let fechaTexto = "9999-12-31";
            let horaTexto = "00:00";
            let esHoy = false;

            if (fechaCruda && typeof fechaCruda === 'string') {
                try {
                    const dateObjeto = new Date(fechaCruda);
                    if (!isNaN(dateObjeto.getTime())) {
                        fechaTexto = dateObjeto.toLocaleDateString('sv-SE', { timeZone: 'America/Santiago' });
                        esHoy = (fechaTexto === hoyChile);
                        
                        horaTexto = dateObjeto.toLocaleTimeString('es-CL', { 
                            hour: '2-digit', 
                            minute: '2-digit', 
                            hour12: false, 
                            timeZone: 'America/Santiago' 
                        });
                    }
                } catch (e) {
                    // Si falla el objeto Date, intentamos romper el String crudo (Ej: 2026-05-26T14:30:00Z)
                    const partes = fechaCruda.split('T');
                    if (partes[0]) {
                        fechaTexto = partes[0];
                        esHoy = (fechaTexto === hoyChile);
                    }
                    if (partes[1]) {
                        horaTexto = partes[1].substring(0, 5);
                    }
                }
            }

            return {
                title: f.Title || f.LinkTitle || "Evento sin título",
                sala: salaReal,
                hora: horaTexto,
                fechaStr: fechaTexto,
                esHoy: esHoy,
                casillaTiempo: fechaCruda
            };
        });

        // Ordenar cronológicamente
        eventosProcesados.sort((a, b) => `${a.fechaStr}T${a.hora}`.localeCompare(`${b.fechaStr}T${b.hora}`));

        res.setHeader('Cache-Control', 'no-shadow, no-store, must-revalidate');
        return res.status(200).json(eventosProcesados);

    } catch (error) {
        console.error("Fallo crítico:", error);
        return res.status(500).json({ error: error.message });
    }
};

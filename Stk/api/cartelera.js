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

        const graphUrl = `https://graph.microsoft.com/v1.0/sites/ugmchile.sharepoint.com:/sites/Calen:/lists/lista%20calen/items?expand=fields&$top=100`;
        const graphRes = await fetch(graphUrl, {
            method: 'GET',
            headers: { 'Authorization': `Bearer ${accessToken}`, 'Accept': 'application/json' }
        });

        if (!graphRes.ok) throw new Error("Microsoft Graph no respondió");
        const graphData = graphRes.json ? await graphRes.json() : JSON.parse(await graphRes.text());
        const rawItems = graphData.value || [];

        const hoyChile = new Date().toLocaleDateString('sv-SE', { timeZone: 'America/Santiago' });

        const eventosProcesados = rawItems.map(item => {
            const f = item.fields || {};
            
            // 🚀 CORRECCIÓN DE UBICACIÓN: Tu lista usa U_G_M_Sala
            let sala = f.U_G_M_Sala || f.Sala || f.Ubicacion || "Por definir";
            
            let fechaCruda = f.EventDate || f.StartDate || "";
            let fechaTexto = "9999-12-31";
            let horaTexto = "00:00";
            let esHoy = false;

            if (fechaCruda && typeof fechaCruda === 'string') {
                const partes = fechaCruda.split('T');
                if (partes[0] && partes[1]) {
                    fechaTexto = partes[0];
                    esHoy = (fechaTexto === hoyChile);
                    
                    // 🚀 CORRECCIÓN DE HORA: Formateamos usando el huso horario local de Santiago
                    try {
                        const dateObjeto = new Date(fechaCruda);
                        if (!isNaN(dateObjeto.getTime())) {
                            horaTexto = dateObjeto.toLocaleTimeString('es-CL', { 
                                hour: '2-digit', 
                                minute: '2-digit', 
                                hour12: false, 
                                timeZone: 'America/Santiago' 
                            });
                        }
                    } catch (e) {
                        const subPartesHora = partes[1].split(':');
                        let horaOriginal = parseInt(subPartesHora[0], 10);
                        let minutos = subPartesHora[1] || "00";
                        if (!isNaN(horaOriginal)) {
                            let horaAjustada = (horaOriginal + 4) % 24;
                            horaTexto = `${String(horaAjustada).padStart(2, '0')}:${minutos}`;
                        }
                    }
                }
            }

            return {
                title: f.Title || f.LinkTitle || "Evento sin título",
                sala: sala,
                hora: horaTexto,
                fechaStr: fechaTexto,
                esHoy: esHoy
            };
        });

        // Entregamos la lista completa limpia al frontend sin filtros duros que la dejen en blanco
        eventosProcesados.sort((a, b) => `${a.fechaStr}T${a.hora}`.localeCompare(`${b.fechaStr}T${b.hora}`));

        res.setHeader('Cache-Control', 'no-shadow, no-store, must-revalidate');
        return res.status(200).json(eventosProcesados);

    } catch (error) {
        console.error(error);
        return res.status(500).json({ error: error.message });
    }
};

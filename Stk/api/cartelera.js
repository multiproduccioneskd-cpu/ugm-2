module.exports = async (req, res) => {
    // Cabeceras CORS obligatorias para que el navegador de la tele no bloquee la petición
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

        if (!tokenRes.ok) {
            const errText = await tokenRes.text();
            throw new Error(`Error Azure Token: ${errText}`);
        }
        
        const tokenData = await tokenRes.json();
        const accessToken = tokenData.access_token;

        // 2. 🚀 LA URL REAL DE TU HISTORIAL: Apunta al sitio 'Calen' y la lista 'lista calen'
        const graphUrl = `https://graph.microsoft.com/v1.0/sites/ugmchile.sharepoint.com:/sites/Calen:/lists/lista%20calen/items?expand=fields&$top=100`;

        const graphRes = await fetch(graphUrl, {
            method: 'GET',
            headers: { 
                'Authorization': `Bearer ${accessToken}`,
                'Accept': 'application/json',
                'Cache-Control': 'no-cache'
            }
        });

        if (!graphRes.ok) {
            const errText = await graphRes.text();
            throw new Error(`Microsoft Graph respondió: ${errText}`);
        }
        
        const graphData = await graphRes.json();
        const rawItems = graphData.value || [];

        // Obtener fecha de hoy y hora en Chile (Santiago) para realizar los cortes cronológicos
        const hoyChile = new Date().toLocaleDateString('sv-SE', { timeZone: 'America/Santiago' });
        const horaChile = new Date().toLocaleTimeString('es-CL', { hour: '2-digit', minute: '2-digit', hour12: false, timeZone: 'America/Santiago' });

        // 3. Mapeo idéntico al de tu Git historial que extrae horas y salas correctas
        const eventosProcesados = rawItems.map(item => {
            const f = item.fields || {};
            
            // Forzar rastreo de ubicación
            let sala = f.Sala || f.Ubicacion || f.Location || f.U_G_M_Sala || "Por definir";
            
            // Rastreador agresivo de fecha de calendario real
            let fechaCruda = f.EventDate || f.StartDate || f.EventDateTime || item.createdDateTime || "";
            let fechaTexto = "9999-12-31";
            let horaTexto = "00:00";
            let esHoy = false;

            if (fechaCruda && typeof fechaCruda === 'string') {
                const partes = fechaCruda.split('T');
                if (partes[0] && partes[1]) {
                    fechaTexto = partes[0];
                    esHoy = (fechaTexto === hoyChile);
                    
                    const subPartesHora = partes[1].split(':');
                    let horaOriginal = parseInt(subPartesHora[0], 10);
                    let minutos = subPartesHora[1] || "00";
                    
                    if (!isNaN(horaOriginal)) {
                        // Sumamos las 4 horas de desfase de la API de Microsoft para la hora local de Chile
                        let horaAjustada = (horaOriginal + 4) % 24;
                        if (parseInt(minutos, 10) > 0 && parseInt(minutos, 10) < 10) minutos = "00";
                        horaTexto = `${String(horaAjustada).padStart(2, '0')}:${minutos}`;
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

        // 4. Separar la data según lo que pida el parámetro ?vista=
        const tipoVista = req.query.vista || 'hoy';
        let resultadoFiltrado = [];

        if (tipoVista === 'semana') {
            resultadoFiltrado = eventosProcesados.filter(ev => !ev.esHoy && ev.fechaStr > hoyChile);
        } else {
            resultadoFiltrado = eventosProcesados.filter(ev => ev.esHoy && ev.hora >= horaChile);
        }

        // Si la de hoy está vacía, le manda los próximos de una para que la tele no quede en blanco
        if (resultadoFiltrado.length === 0 && tipoVista === 'hoy') {
            resultadoFiltrado = eventosProcesados.filter(ev => !ev.esHoy && ev.fechaStr > hoyChile);
        }

        // Ordenar cronológicamente
        resultadoFiltrado.sort((a, b) => `${a.fechaStr}T${a.hora}`.localeCompare(`${b.fechaStr}T${b.hora}`));

        res.setHeader('Cache-Control', 'no-shadow, no-store, must-revalidate');
        return res.status(200).json(resultadoFiltrado);

    } catch (error) {
        console.error("Fallo crítico backend:", error);
        return res.status(500).json({ error: error.message });
    }
};

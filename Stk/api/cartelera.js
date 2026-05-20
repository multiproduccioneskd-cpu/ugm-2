export default async function handler(req, res) {
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
    res.setHeader('Access-Control-Allow-Methods', 'GET, OPTIONS');

    if (req.method === 'OPTIONS') return res.status(200).end();

    try {
        const { TENANT_ID, CLIENT_ID, CLIENT_SECRET } = process.env;

        const tokenUrl = `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`;
        const params = new URLSearchParams({
            client_id: CLIENT_ID,
            scope: 'https://graph.microsoft.com/.default',
            client_secret: CLIENT_SECRET,
            grant_type: 'client_credentials'
        });

        const tokenRes = await fetch(tokenUrl, {
            method: 'POST',
            body: params,
            headers: { 'Content-Type': 'application/x-www-form-urlencoded' }
        });
        
        if (!tokenRes.ok) throw new Error("Error obteniendo Token de Azure");
        const tokenData = await tokenRes.json();
        const accessToken = tokenData.access_token;

        // Consultamos la lista expandiendo todas las propiedades ocultas
        const graphUrl = `https://graph.microsoft.com/v1.0/sites/ugmchile.sharepoint.com:/sites/Calen:/lists/lista%20calen/items?expand=fields&$top=100`;
        const graphRes = await fetch(graphUrl, {
            method: 'GET',
            headers: { 'Authorization': `Bearer ${accessToken}`, 'Accept': 'application/json' }
        });

        if (!graphRes.ok) throw new Error("Microsoft Graph no respondió");
        const graphData = await graphRes.json();
        const rawItems = graphData.value || [];

        // Obtener fecha de hoy en Chile (YYYY-MM-DD)
        const hoyChile = new Date().toLocaleDateString('sv-SE', { timeZone: 'America/Santiago' });
        const horaChile = new Date().toLocaleTimeString('es-CL', { hour: '2-digit', minute: '2-digit', hour12: false, timeZone: 'America/Santiago' });

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
                        // Sumamos las 4 horas de desfase de la API de Microsoft
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

        // Separar de forma dura en el backend según lo que pida el frontend
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

        return res.status(200).json(resultadoFiltrado);

    } catch (error) {
        return res.status(500).json({ error: error.message });
    }
}

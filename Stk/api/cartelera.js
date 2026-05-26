module.exports = async (req, res) => {
    // Cabeceras CORS limpias para la tele
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

        // 2. 🚀 LA URL REAL CON EL EXPAND INCLUIDO: Apunta a tu sitio 'Calen' abriendo la llave de los 'fields'
        const graphUrl = `https://graph.microsoft.com/v1.0/sites/ugmchile.sharepoint.com:/sites/Calen:/lists/lista%20calen/items?expand=fields&$top=100`;

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

        // Obtener la fecha de hoy real en formato Santiago Chile
        const hoyChile = new Date().toLocaleDateString('sv-SE', { timeZone: 'America/Santiago' });

        // 3. Mapeo ultra-tolerante para extraer salas, horas y fechas verdaderas
        const eventosProcesados = rawItems.map(item => {
            const f = item.fields || {};
            
            // Extrae la sala real desde tu campo U_G_M_Sala de SharePoint
            let salaReal = f.U_G_M_Sala || f.Sala || f.Ubicacion || "Por definir";
            
            // Busca la fecha de calendario en EventDate o StartDate
            let fechaCruda = f.EventDate || f.StartDate || f.EventDateTime || "";
            let fechaTexto = "9999-12-31";
            let horaTexto = "00:00";
            let esHoy = false;

            if (fechaCruda && typeof fechaCruda === 'string') {
                try {
                    const dateObjeto = new Date(fechaCruda);
                    if (!isNaN(dateObjeto.getTime())) {
                        // Formateamos la fecha exacta para Chile (YYYY-MM-DD)
                        fechaTexto = dateObjeto.toLocaleDateString('sv-SE', { timeZone: 'America/Santiago' });
                        esHoy = (fechaTexto === hoyChile);
                        
                        // Formateamos la hora exacta local de Santiago Chile (HH:MM)
                        horaTexto = dateObjeto.toLocaleTimeString('es-CL', { 
                            hour: '2-digit', 
                            minute: '2-digit', 
                            hour12: false, 
                            timeZone: 'America/Santiago' 
                        });
                    }
                } catch (e) {
                    // Respaldo por string plano si el objeto Date falla
                    const partes = fechaCruda.split('T');
                    if (partes[0]) {
                        fechaTexto = partes[0];
                        esHoy = (fechaTexto === hoyChile);
                    }
                }
            }

            return {
                title: f.Title || f.LinkTitle || "Evento sin título",
                sala: salaReal,
                hora: horaTexto,
                fechaStr: fechaTexto,
                esHoy: esHoy
            };
        });

        // 4. Ordenar cronológicamente para que se liste bien la cascada en la tele
        eventosProcesados.sort((a, b) => `${a.fechaStr}T${a.hora}`.localeCompare(`${b.fechaStr}T${b.hora}`));

        res.setHeader('Cache-Control', 'no-shadow, no-store, must-revalidate');
        return res.status(200).json(eventosProcesados);

    } catch (error) {
        console.error("Fallo crítico en el procesador:", error);
        return res.status(500).json({ error: error.message });
    }
};

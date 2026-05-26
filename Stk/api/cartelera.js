module.exports = async (req, res) => {
    // Cabeceras CORS duras para que el visor de la tele lea sin bloqueos
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
    res.setHeader('Access-Control-Allow-Methods', 'GET, OPTIONS');

    if (req.method === 'OPTIONS') {
        res.status(200).end();
        return;
    }

    try {
        const { TENANT_ID, CLIENT_ID, CLIENT_SECRET } = process.env;

        // 1. Obtener Token de Azure
        const tokenUrl = `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`;
        const params = new URLSearchParams({
            client_id: CLIENT_ID,
            scope: 'https://graph.microsoft.com/.default',
            client_secret: CLIENT_SECRET,
            grant_type: 'client_credentials'
        });

        const tokenRes = await fetch(tokenUrl, {
            method: 'POST',
            body: params.toString(),
            headers: { 'Content-Type': 'application/x-www-form-urlencoded' }
        });

        if (!tokenRes.ok) {
            const errText = await tokenRes.text();
            throw new Error(`Error Azure Token: ${errText}`);
        }
        
        const tokenData = await tokenRes.json();
        const accessToken = tokenData.access_token;

        // 2. Traer datos desde SharePoint apuntando a la lista 'lista calen'
        const graphUrl = `https://graph.microsoft.com/v1.0/sites/ugmchile.sharepoint.com:/sites/Calen:/lists/lista%20calen/items?expand=fields&$top=100`;

        const graphRes = await fetch(graphUrl, {
            method: 'GET',
            headers: { 
                'Authorization': `Bearer ${accessToken}`,
                'Accept': 'application/json'
            }
        });

        if (!graphRes.ok) {
            const errText = await graphRes.text();
            throw new Error(`Microsoft Graph respondió: ${errText}`);
        }
        
        const graphData = await graphRes.json();
        const rawItems = graphData.value || [];

        // 3. Capturar tiempos en Santiago de Chile
        const hoyChile = new Date().toLocaleDateString('sv-SE', { timeZone: 'America/Santiago' }); 

        // 4. Mapear y limpiar la data de origen exactamente como lo hacía el tuyo
        const eventosProcesados = rawItems.map(item => {
            const f = item.fields || {};
            
            // Buscador secuencial de ubicación/sala
            let sala = "Por definir";
            const posiblesColumnasSala = ["Sala", "Ubicacion", "Location", "U_G_M_Sala", "Ubicaci_x00f3_n"];
            for (let col of posiblesColumnasSala) {
                if (f[col]) { 
                    sala = f[col]; 
                    break; 
                }
            }

            let fechaTexto = "9999-12-31";
            let horaTexto = "00:00";
            let esHoy = false;

            // 🚀 LA FÓRMULA HISTÓRICA DEL DESFASE +4
            let fechaCruda = f.EventDate || f.StartDate || f.EventDateTime || item.createdDateTime || "";
            if (fechaCruda && typeof fechaCruda === 'string') {
                const partes = fechaCruda.split('T');
                if (partes[0] && partes[1]) {
                    fechaTexto = partes[0];
                    esHoy = (fechaTexto === hoyChile);
                    
                    const subPartesHora = partes[1].split(':');
                    let horaOriginal = parseInt(subPartesHora[0], 10);
                    let minutos = subPartesHora[1] || "00";
                    
                    if (!isNaN(horaOriginal)) {
                        // Sumamos las 4 horas de desfase UTC-4 para la hora nacional
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
                esHoy: esHoy,
                id: item.id,
                fields: f // Mandamos el objeto completo por si el HTML procesa algo extra por fuera
            };
        });

        // Ordenamiento cronológico antes de enviar el paquete completo
        eventosProcesados.sort((a, b) => `${a.fechaStr}T${a.hora}`.localeCompare(`${b.fechaStr}T${b.hora}`));

        res.setHeader('Cache-Control', 'no-shadow, no-store, must-revalidate');
        return res.status(200).json(eventosProcesados);

    } catch (error) {
        console.error("Fallo crítico backend:", error);
        return res.status(500).json({ 
            error: "Error interno del servidor backend", 
            mensaje: error.message 
        });
    }
};

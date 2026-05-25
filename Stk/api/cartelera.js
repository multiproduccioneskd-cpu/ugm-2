       const tokenData = await tokenRes.json();
        const accessToken = tokenData.access_token;

        const graphUrl = `https://graph.microsoft.com/v1.0/sites/ugmchile.sharepoint.com:/sites/Calen:/lists/lista%20calen/items?expand=fields&$top=100`;
        // Consultamos la lista de SharePoint
        const graphUrl = `https://graph.microsoft.com/v1.0/sites/ugmchile.sharepoint.com:/sites/Calen:/lists/lista%20calen/items?expand=fields&$top=5`;
        const graphRes = await fetch(graphUrl, {
            method: 'GET',
            headers: { 'Authorization': `Bearer ${accessToken}`, 'Accept': 'application/json' }
@@ -36,48 +37,15 @@ module.exports = async function handler(req, res) {
        const graphData = await graphRes.json();
        const rawItems = graphData.value || [];

        const eventosNormalizados = rawItems.map(item => {
            const f = item.fields || {};
            
            let salaReal = f.Sala || f.Ubicacion || f.Location || f.U_G_M_Sala || "Por definir";
            if (salaReal === "Por definir") {
                for (let key in f) {
                    if (typeof f[key] === 'string' && (f[key].includes("Auditorio") || f[key].includes("Gimnasio") || f[key].includes("Sala") || f[key].includes("SALÓN"))) {
                        salaReal = f[key];
                        break;
                    }
                }
            }
        // HACK DE INSPECCIÓN: Mandamos las entrañas completas del primer ítem para pillar los campos ocultos
        if (rawItems.length > 0) {
            return res.status(200).json({
                _INSPECCION_RAIZ: Object.keys(rawItems[0]),
                _INSPECCION_FIELDS: rawItems[0].fields || {}
            });
        }

            // BUSCADOR EN DETALLE DE LA FECHA REAL DEL EVENTO
            // Priorizamos las propiedades nativas de eventos de SharePoint que oculta Graph
            let casillaTiempo = f.EventDate || f.OData__StartDate || f.StartDate || f.EventDateTime || f.Fecha || "";
            
            // Si no encontró ninguna de las anteriores, buscamos cualquier llave que tenga "Start" o "Date" pero ignorando los metadatos de modificación del archivo
            if (!casillaTiempo) {
                for (let key in f) {
                    if (key.toLowerCase().includes('start') || (key.toLowerCase().includes('date') && !key.toLowerCase().includes('modify') && !key.toLowerCase().includes('author'))) {
                        if (f[key] && typeof f[key] === 'string' && f[key].includes('T')) {
                            casillaTiempo = f[key];
                            break;
                        }
                    }
                }
            }

            // Caída de emergencia total si de verdad la lista no tiene columnas de calendario
            if (!casillaTiempo) {
                casillaTiempo = item.createdDateTime || f.Modified || "";
            }

            return {
                title: f.Title || f.LinkTitle || "Evento sin título",
                sala: salaReal,
                casillaTiempo: casillaTiempo
            };
        });

        return res.status(200).json(eventosNormalizados);
        return res.status(200).json({ mensaje: "Lista vacía" });

    } catch (error) {
        return res.status(500).json({ error: true, mensaje: error.message });

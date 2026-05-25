const { AuthType, pnpRepository } = require('@pnp/sp-common-js'); // O la librería de conexión que uses en tu proyecto

export default async function handler(req, res) {
    // Configurar cabeceras CORS básicas para que la tele pueda leer la API sin bloqueos
    res.setHeader('Access-Control-Allow-Credentials', true);
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.setHeader('Access-Control-Allow-Methods', 'GET,OPTIONS');
    res.setHeader('Access-Control-Allow-Headers', 'X-CSRF-Token, X-Requested-With, Accept, Accept-Version, Content-Length, Content-MD5, Content-Type, Date, X-Api-Version');

    if (req.method === 'OPTIONS') {
        res.status(200).end();
        return;
    }

    try {
        // 1. AQUÍ TU API CONECTA CON LA LISTA DE SHAREPOINT (Simulando la consulta a la SDK)
        // Revisa que el llamado traiga todos los campos ('fields') de la lista de eventos UGM.
        const listItems = await obtenerItemsDesdeSharePoint(); 

        // 2. EL GOLAZO: Procesamos el arreglo mapeando los campos uno a uno hacia el frontend
        const eventosProcesados = listItems.map(item => {
            // Aseguramos la lectura del campo Destinatario por si viene con mayúscula o minúscula
            const rawDestinatario = item.fields.Destinatario || item.fields.destinatario || null;

            return {
                title: item.fields.Title || item.fields.title || "Evento sin título",
                sala: item.fields.Sala || item.fields.sala || "Por definir",
                casillaTiempo: item.fields.Fecha || item.fields.casillaTiempo || item.fields.EventDate || null,
                
                // Enviamos el dato al index.html de forma limpia
                Destinatario: rawDestinatario
            };
        });

        // 3. RESPUESTA: Entregamos el JSON estructurado al frontend con no-cache activado
        res.setHeader('Cache-Control', 'no-shadow, no-store, must-revalidate');
        res.status(200).json(eventosProcesados);

    } catch (error) {
        console.error("Error en el backend de cartelera:", error);
        res.status(500).json({ error: "Error interno al conectar con SharePoint", detalle: error.message });
    }
}

// 👁️ NOTA DE CONTROL: Esta es la función interna que va a buscar la data a Microsoft. 
// Mantén la lógica de autenticación (tenantId, clientId, clientSecret) que ya tenías funcionando de pana.
async function obtenerItemsDesdeSharePoint() {
    // Aquí corre tu código actual de conexión (ej: axios a Microsoft Graph o @pnp/sp)
    // Asegúrate de que el query traiga la nueva columna 'Destinatario'.
    // Retorna el arreglo de filas (items) directo de la lista de SharePoint.
}

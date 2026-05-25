// 1. MANTÉN AQUÍ TUS CONFIGURACIONES ORIGINALES
// Deja intactas todas las librerías, los tokens, clientSecret y accesos de tu código viejo.

export default async function handler(req, res) {
    // Tus cabeceras CORS que ya tenías...
    res.setHeader('Access-Control-Allow-Origin', '*');

    try {
        // 2. MANTÉN TU CONSULTA ORIGINAL A SHAREPOINT
        // La línea exacta con la que traías la data antes (ej: await tuFuncionDeSharePoint())
        const listItems = await tuFuncionOriginalDeAntes(); 

        // 3. AQUÍ HACEMOS EL MAPEO REPARADO PARA EL FRONTEND
        const eventosProcesados = listItems.map(item => {
            return {
                // Mantén los nombres exactos de los campos como los tenías en tu código del miércoles:
                title: item.fields.Title || item.fields.title,
                sala: item.fields.Sala || item.fields.sala,
                casillaTiempo: item.fields.Fecha || item.fields.casillaTiempo,
                
                // 🚀 EL GOLAZO: Solo sumamos esto al final del objeto original:
                Destinatario: item.fields.Destinatario || item.fields.destinatario || null
            };
        });

        // 4. MANTÉN TU RESPUESTA ORIGINAL
        res.status(200).json(eventosProcesados);

    } catch (error) {
        console.error(error);
        res.status(500).json({ error: "Error" });
    }
}

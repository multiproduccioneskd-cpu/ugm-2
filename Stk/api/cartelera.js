import { ConfidentialClientApplication } from "@azure/msal-node";
import axios from "axios";

// 🛠️ CONFIGURACIÓN DE CREDENCIALES (Usa las variables de entorno de tu Vercel)
const msalConfig = {
    auth: {
        clientId: process.env.CLIENT_ID || process.env.NEXT_PUBLIC_CLIENT_ID,
        authority: `https://login.microsoftonline.com/${process.env.TENANT_ID}`,
        clientSecret: process.env.CLIENT_SECRET,
    }
};

const tokenRequest = {
    scopes: ["https://graph.microsoft.com/.default"],
};

const cca = new ConfidentialClientApplication(msalConfig);

async function getAccessToken() {
    const response = await cca.acquireTokenByClientCredential(tokenRequest);
    return response.accessToken;
}

export default async function handler(req, res) {
    // Cabeceras CORS obligatorias para la tele de la U
    res.setHeader('Access-Control-Allow-Credentials', true);
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.setHeader('Access-Control-Allow-Methods', 'GET,OPTIONS');
    res.setHeader('Access-Control-Allow-Headers', 'X-CSRF-Token, X-Requested-With, Accept, Accept-Version, Content-Length, Content-MD5, Content-Type, Date, X-Api-Version');

    if (req.method === 'OPTIONS') {
        res.status(200).end();
        return;
    }

    try {
        const token = await getAccessToken();
        
        // Llamado limpio a Microsoft Graph usando tus variables de entorno del Site y la Lista
        const url = `https://graph.microsoft.com/v1.0/sites/${process.env.SHAREPOINT_SITE_ID}/lists/${process.env.SHAREPOINT_LIST_ID}/items?expand=fields`;
        
        const response = await axios.get(url, {
            headers: {
                Authorization: `Bearer ${token}`,
                Accept: "application/json",
            }
        });

        const items = response.data.value || [];

        // 🚀 EL MAPEO ORIGINAL REPARADO: Aquí procesamos la lista
        const eventosProcesados = items.map(item => {
            return {
                title: item.fields.Title || item.fields.title || "Evento sin título",
                sala: item.fields.Sala || item.fields.sala || "Por definir",
                casillaTiempo: item.fields.Fecha || item.fields.casillaTiempo || null,
                
                // Aquí agarra la columna tipo Elección sin que se vaya a negro
                Destinatario: item.fields.Destinatario || item.fields.destinatario || null
            };
        });

        // Forzamos a Vercel a no guardar caché vieja
        res.setHeader('Cache-Control', 'no-shadow, no-store, must-revalidate');
        res.status(200).json(eventosProcesados);

    } catch (error) {
        console.error("Error en cartelera API:", error.response ? error.response.data : error.message);
        res.status(500).json({ error: "Error de conexión con SharePoint", detalle: error.message });
    }
}

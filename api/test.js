export default function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET');
   
  return res.json({ 
    status: "online",
    service: "API UniBiotech Excel Export",
    version: "1.0.0",
    timestamp: new Date().toISOString(),
    endpoints: [
      "POST /api/export/fabricados - Exportar fabricados",
      "POST /api/export/transportados - Exportar transportados",
      "POST /api/export/congelados - Exportar congelados",
      "POST /api/export/all-in-one - Exportar completo"
    ],
    docs: "Envie um POST com { data: [...] } para exportar"
  });
}

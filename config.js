window.APP_CONFIG = {
  timezone: "America/Guayaquil",

  msal: {
    clientId: "PEGA_AQUI_TU_CLIENT_ID",
    tenantId: "PEGA_AQUI_TU_TENANT_ID",
    // Scopes delegados recomendados para SharePoint/Excel:
    scopes: ["User.Read", "Files.ReadWrite.All"]
  },

  shareLinks: {
    // Archivo lista de bloques (tu link):
    blocks: "https://malimaec-my.sharepoint.com/:x:/g/personal/analistabi_malima_com/IQDomkop5tidRZDuKVqeMqCXAQyn8rpZgQi62-qeMqvAGtA?e=Adbu4H",
    // Archivo base con tbl_muestras (tu link):
    data: "https://malimaec-my.sharepoint.com/:x:/g/personal/analistabi_malima_com/IQChxrCEhJT6TYzm6DpTlIt0ASdRW7QAJ9q-Mx9-3TjxdDc?e=x6P7xR"
  },

  // Hoja donde están los bloques (usedRange)
  blocksSheetName: "bloques",

  // Encabezados posibles (case-insensitive). Ajusta si tus columnas se llaman distinto.
  blocksHeaderCandidates: {
    bloque: ["bloque", "BLOQUE", "Bloque"],
    estado: ["estado", "ESTADO", "Estado"]
  },

  // Tablas en tu Excel "data"
  tables: {
    muestras: "tbl_muestras",
    users: "tbl_users"
  },

  // Defaults de N por modo (editable en UI)
  defaults: {
    nAltura: 16,
    nCompleto: 120
  },

  // Validaciones Gypsophila (ajusta a tu realidad)
  validation: {
    altura_cm: { min: 5, max: 200 },
    estructura_cm: { min: 1, max: 200 },
    diametro_mm: { min: 0.5, max: 30 }
  },

  // Texto exacto para estado activo
  estadoActivo: "ACTIVO",

  // Mensajes
  ui: {
    appTitle: "Formulario de Muestreo"
  }
};

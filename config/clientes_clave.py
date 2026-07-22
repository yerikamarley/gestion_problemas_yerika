"""Catálogo y alias de los clientes incluidos en el seguimiento especial."""

ASOBANCARIA = [
    "Asobancaria", "Bancolombia", "Bancoomeva", "BBVA", "Banco de Bogotá",
    "Banco Caja Social", "Davivienda", "Banco Falabella", "Mibanco",
    "Banco Mundo Mujer", "Banco de Occidente", "Banco Popular", "Porvenir",
    "RCI", "Sufi", "Tuya", "Banco AV Villas",
]

CONFECAMARAS = [
    "Cámara de Comercio de Bogotá", "Cámara de Comercio de Facatativá",
    "Cámara de Comercio de Girardot, Alto Magdalena y Tequendama",
    "Cámara de Comercio de Tunja", "Cámara de Comercio de Duitama",
    "Cámara de Comercio de Sogamoso", "Cámara de Comercio de Barranquilla",
    "Cámara de Comercio de Cartagena",
    "Cámara de Comercio de Santa Marta para el Magdalena",
    "Cámara de Comercio de Valledupar para el Valle del Río Cesar",
    "Cámara de Comercio de Aguachica", "Cámara de Comercio de la Guajira",
    "Cámara de Comercio de Montería", "Cámara de Comercio de Sincelejo",
    "Cámara de Comercio de Magangué",
    "Cámara de Comercio de San Andrés, Providencia y Santa Catalina",
    "Cámara de Comercio de Medellín para Antioquia",
    "Cámara de Comercio del Aburrá Sur", "Cámara de Comercio del Oriente Antioqueño",
    "Cámara de Comercio del Urabá",
    "Cámara de Comercio del Magdalena Medio y Nordeste Antioqueño",
    "Cámara de Comercio del Chocó", "Cámara de Comercio de Cali",
    "Cámara de Comercio de Buenaventura", "Cámara de Comercio de Buga",
    "Cámara de Comercio de Cartago", "Cámara de Comercio de Palmira",
    "Cámara de Comercio de Tuluá", "Cámara de Comercio del Cauca",
    "Cámara de Comercio de Pasto", "Cámara de Comercio de Ipiales",
    "Cámara de Comercio de Tumaco", "Cámara de Comercio de Manizales por Caldas",
    "Cámara de Comercio de La Dorada, Puerto Boyacá, Puerto Salgar y Oriente de Caldas",
    "Cámara de Comercio de Pereira", "Cámara de Comercio de Dosquebradas",
    "Cámara de Comercio de Santa Rosa de Cabal",
    "Cámara de Comercio de Armenia y del Quindío",
    "Cámara de Comercio de Bucaramanga", "Cámara de Comercio de Barrancabermeja",
    "Cámara de Comercio de Cúcuta", "Cámara de Comercio de Ocaña",
    "Cámara de Comercio de Pamplona", "Cámara de Comercio de Ibagué",
    "Cámara de Comercio de Honda, Guaduas y Norte del Tolima",
    "Cámara de Comercio del Sur y Oriente del Tolima", "Cámara de Comercio del Huila",
    "Cámara de Comercio de Villavicencio", "Cámara de Comercio de Arauca",
    "Cámara de Comercio del Piedemonte Araucano", "Cámara de Comercio de Casanare",
    "Cámara de Comercio de Florencia para el Caquetá",
    "Cámara de Comercio del Putumayo", "Cámara de Comercio del Amazonas",
    "Cámara de Comercio de San José del Guaviare", "Cámara de Comercio del Guainía",
    "Cámara de Comercio del Vaupés",
]

OTROS_CLIENTES_CLAVE = [
    "Claro", "Seguros del Estado", "Colpensiones", "Enterritorio",
    "Empresa de Energía de Pereira", "REN Consultores",
]

GRUPOS_CLIENTES_CLAVE = {
    "Asobancaria": ASOBANCARIA,
    "Confecámaras": CONFECAMARAS,
    "Otros clientes clave": OTROS_CLIENTES_CLAVE,
}

CLIENTES_CLAVE = [
    cliente
    for clientes in GRUPOS_CLIENTES_CLAVE.values()
    for cliente in clientes
]

# Los nombres oficiales siempre son alias válidos. Las variantes cubren valores
# frecuentes encontrados en las fuentes sin cambiar el nombre mostrado al usuario.
CLIENTES_CLAVE_ALIASES = {cliente: [cliente] for cliente in CLIENTES_CLAVE}
CLIENTES_CLAVE_ALIASES.update(
    {
        "Bancolombia": ["Bancolombia", "Bancolombia S.A"],
        "Bancoomeva": ["Bancoomeva", "Bancoomeva S.A"],
        "Banco Caja Social": ["Banco Caja Social", "Caja Social"],
        "Davivienda": ["Davivienda", "Banco Davivienda", "Banco Davivienda S.A."],
        "Mibanco": ["Mibanco", "Mibanco S.A."],
        "RCI": [
            "RCI",
            "RCI Colombia",
            "RCI COLOMBIA S.A COMPAÑÍA DE FINANCIAMIENTO",
        ],
        "Sufi": ["Sufi", "SUFI BANCOLOMBIA"],
        "Tuya": ["Tuya", "Tuya S.A"],
        "Banco AV Villas": ["Banco AV Villas", "AV Villas"],
        "Claro": ["Claro", "Comcel"],
        "Colpensiones": ["Colpensiones", "Colpen"],
    }
)


/**
 * niva — Google Apps Script para emails automáticos
 *
 * INSTRUCCIONES DE INSTALACIÓN:
 * 1. Abrí tu Google Form en modo edición
 * 2. Click en los 3 puntos (⋮) → "Editor de secuencias de comandos"
 * 3. Borrá todo el contenido y pegá este archivo completo
 * 4. Guardá (Ctrl+S)
 * 5. Andá a "Activadores" (ícono de reloj) → "Agregar activador"
 *    - Función: onFormSubmit
 *    - Evento: Al enviar formulario
 * 6. Autorizá los permisos
 *
 * IMPORTANTE — CONFIGURACIÓN ADICIONAL:
 *
 * A) DESACTIVAR notificación por defecto del formulario:
 *    Google Form → Respuestas → clic en los 3 puntos (⋮) →
 *    desmarcar "Recibir notificaciones por correo electrónico de nuevas respuestas"
 *    (Esto evita el email viejo con "montevideo" en el footer)
 *
 * B) CONFIGURAR "Enviar como" en Gmail para que aparezca hola@niva.uy:
 *    1. Gmail → Configuración (⚙️) → Ver todos los ajustes → Cuentas e importación
 *    2. En "Enviar como:" → "Añadir otra dirección de correo electrónico"
 *    3. Nombre: niva | Email: hola@niva.uy
 *    4. Desmarcar "Tratar como alias"
 *    5. Servidor SMTP: smtp.gmail.com | Puerto: 587
 *       (Si usás Cloudflare Email Routing, Gmail lo maneja automáticamente
 *        — solo confirmá el código de verificación que llega a hola@niva.uy)
 *    6. Te llega un email de confirmación a hola@niva.uy → hacé click en el link
 *    7. Una vez confirmado, cambiá aquí USE_GMAIL_SEND_AS = true
 */

// ===== CONFIGURACIÓN =====
const NIVA_EMAIL = 'hola@niva.uy';
const NIVA_REPLY_TO = 'hola@niva.uy';
const NIVA_ADMIN = 'alejandrabotta@gmail.com'; // para notificaciones internas

// ID del spreadsheet de respuestas (sacalo de la URL: docs.google.com/spreadsheets/d/ESTE_ID/edit)
const SPREADSHEET_ID = '1SP2yRYcnpR4bxY1nOBb3dKKAypWSLle7e46hdyQUENw';

// Colores de los discos por suplemento
const DISC_COLORS = {
  'Complejo B': '#B8C1B6',
  'Magnesio': '#C9C5BE',
  'Magnesio Citrato': '#C9C5BE',
  'Maca': '#D4C4A8',
  'Ashwagandha': '#9DA89A',
  'Omega 3': '#C5BFAE',
  'Zinc + Vitamina C': '#E0D4B8',
  'Probióticos': '#E9E6E1',
  'Vitamina D + Calcio': '#D6CEB8',
  'Colágeno Hidrolizado': '#DCC9B8',
  'Astaxantina': '#C7A89A',
  'Proteína en Polvo': '#B3B09B',
  'Creatina Monohidrato': '#A8B0A3'
};

const DISC_LABELS = {
  'Complejo B': 'B',
  'Magnesio': 'Mg',
  'Magnesio Citrato': 'Mg',
  'Maca': 'Mc',
  'Ashwagandha': 'Ah',
  'Omega 3': 'O3',
  'Zinc + Vitamina C': 'Zn',
  'Probióticos': 'Pr',
  'Vitamina D + Calcio': 'D3',
  'Colágeno Hidrolizado': 'Co',
  'Astaxantina': 'Ax',
  'Proteína en Polvo': 'Pt',
  'Creatina Monohidrato': 'Cr'
};

const FORMAS = {
  'Complejo B': 'comprimido',
  'Magnesio': 'comprimido',
  'Magnesio Citrato': 'comprimido',
  'Maca': 'cápsula',
  'Ashwagandha': 'cápsula',
  'Omega 3': 'cápsula blanda',
  'Zinc + Vitamina C': 'comprimido',
  'Probióticos': 'cápsula',
  'Vitamina D + Calcio': 'comprimido',
  'Colágeno Hidrolizado': 'polvo',
  'Astaxantina': 'cápsula blanda',
  'Proteína en Polvo': 'polvo',
  'Creatina Monohidrato': 'polvo'
};

// Forma abreviada para mostrar en el disco (sin emojis)
const FORMA_ABREV = {
  'comprimido': 'tab',
  'cápsula': 'cáp',
  'cápsula blanda': 'soft',
  'polvo': 'pwd'
};

// Marca de cada suplemento
const MARCAS = {
  'Complejo B': 'ICU-VITA',
  'Magnesio': 'Qualivits',
  'Magnesio Citrato': 'Qualivits',
  'Maca': 'Qualivits',
  'Ashwagandha': 'Naturaleza Pura',
  'Omega 3': 'Roemmers',
  'Zinc + Vitamina C': 'Vitaminway',
  'Probióticos': 'Biótica Pro',
  'Vitamina D + Calcio': 'Simple (Bagó)',
  'Colágeno Hidrolizado': 'Fix1',
  'Astaxantina': 'Astaxan (Megalabs)',
  'Proteína en Polvo': 'USN',
  'Creatina Monohidrato': 'Edatir (Sylab)'
};

// Formato y duracion de cada suplemento
const FORMATOS = {
  'Complejo B': 'Frasco x 40 comprimidos · Para 40 días',
  'Magnesio': 'Frasco x 100 comprimidos · Para 100 días',
  'Magnesio Citrato': 'Frasco x 100 comprimidos · Para 100 días',
  'Maca': 'Frasco x 100 cápsulas · Para 33 días',
  'Ashwagandha': 'Frasco x 30 cápsulas · Para 30 días',
  'Omega 3': 'Frasco x 100 cápsulas · Para 100 días',
  'Zinc + Vitamina C': 'Frasco x 30 cápsulas · Para 30 días',
  'Probióticos': 'Frasco x 30 cápsulas · Para 30 días',
  'Vitamina D + Calcio': 'Frasco x 60 gummies · Para 30 días',
  'Colágeno Hidrolizado': 'Polvo x 200 g · Para 30 días',
  'Astaxantina': 'Frasco x 30 cápsulas · Para 30 días',
  'Proteína en Polvo': 'Bolsa x 908 g · Para 30 días',
  'Creatina Monohidrato': 'Polvo x 300 g · Para 60 días'
};

// Descripcion breve de cada suplemento (que es)
const QUE_ES = {
  'Complejo B': 'Grupo de 8 vitaminas esenciales que tu cuerpo no produce solo',
  'Magnesio': 'Mineral esencial presente en frutos secos, legumbres y verduras de hoja',
  'Magnesio Citrato': 'Mineral esencial presente en frutos secos, legumbres y verduras de hoja',
  'Maca': 'Raíz andina usada hace miles de años como energizante natural',
  'Ashwagandha': 'Planta adaptógena de la medicina ayurvédica, usada hace más de 3000 años',
  'Omega 3': 'Ácido graso esencial que se obtiene del pescado azul',
  'Zinc + Vitamina C': 'Dos nutrientes clave para tu sistema inmune, presentes en cítricos y carnes',
  'Probióticos': 'Bacterias buenas que equilibran tu flora intestinal',
  'Vitamina D + Calcio': 'Vitamina que se produce con el sol y mineral esencial para tus huesos',
  'Colágeno Hidrolizado': 'Proteína natural que tu cuerpo produce cada vez menos con la edad',
  'Astaxantina': 'Antioxidante natural extraído de microalgas, lo que le da color al salmón',
  'Proteína en Polvo': 'Proteína de suero de leche, la más estudiada para recuperación muscular',
  'Creatina Monohidrato': 'Compuesto natural presente en la carne, el suplemento más estudiado del mundo'
};

// Mapeo nombre → ID del catálogo (para deep links en emails)
const NOMBRE_A_ID = {
  'Complejo B': 'complejo_b',
  'Magnesio': 'magnesio',
  'Magnesio Citrato': 'magnesio',
  'Maca': 'maca',
  'Ashwagandha': 'ashwagandha',
  'Omega 3': 'omega3',
  'Zinc + Vitamina C': 'zinc_vitc',
  'Probióticos': 'probioticos',
  'Vitamina D + Calcio': 'vitamina_d_calcio',
  'Colágeno Hidrolizado': 'colageno',
  'Astaxantina': 'astaxantina',
  'Proteína en Polvo': 'proteina',
  'Creatina Monohidrato': 'creatina'
};

// Mapeo objetivo display → key del catálogo
const OBJETIVO_A_KEY = {
  'Energía y Vitalidad': 'energia',
  'Estrés y Sueño': 'estres',
  'Inmunidad': 'inmunidad',
  'Piel y Pelo': 'piel',
  'Fitness y Rendimiento': 'fitness',
  'energía': 'energia',
  'estrés': 'estres',
  'inmunidad': 'inmunidad',
  'piel': 'piel',
  'fitness': 'fitness',
  'bienestar': 'energia'
};

const CUANDO_TOMAR = {
  'Complejo B': 'Mañana, con el desayuno — da energía, evitar de noche',
  'Magnesio': 'Noche, con la cena o antes de dormir — relaja y ayuda al descanso',
  'Magnesio Citrato': 'Noche, con la cena o antes de dormir — relaja y ayuda al descanso',
  'Maca': 'Mañana, con el desayuno — energizante, empezar con 1 cápsula',
  'Ashwagandha': 'Noche, con la cena — reduce cortisol, mejora el sueño',
  'Omega 3': 'Con la comida más grasa del día — se absorbe mucho mejor con grasa',
  'Zinc + Vitamina C': 'Mañana, en ayunas o desayuno liviano — separar de calcio y magnesio',
  'Probióticos': 'Mañana, en ayunas 20 min antes de comer — llegan más bacterias vivas',
  'Vitamina D + Calcio': 'Mediodía, con el almuerzo — liposoluble, necesita grasa',
  'Colágeno Hidrolizado': 'Mañana, en ayunas — disolver en agua, jugo o café',
  'Astaxantina': 'Con almuerzo o cena — liposoluble, necesita grasa',
  'Proteína en Polvo': 'Post-entreno (dentro de 1 hora) o como snack',
  'Creatina Monohidrato': 'Cualquier momento, todos los días — la consistencia es lo que importa'
};

const CIENCIA = {
  'Complejo B': {
    porque: {
      'energia': 'Convierte lo que comes en energia real. Si sentis cansancio o fatiga durante el dia, las vitaminas B son la base para que tu cuerpo produzca energia de forma eficiente.',
      'estres': 'Tu sistema nervioso necesita vitaminas B para funcionar bien bajo presion. Ayuda a reducir el cansancio mental y mejorar tu estado de animo.',
      'default': 'Las vitaminas B son esenciales para tu energia y tu sistema nervioso. Van a ayudarte a sentirte con mas vitalidad dia a dia.'
    },
    tiempo: '2-4 semanas de uso consistente',
    evidencia: 'Fuerte',
    papers: 'Meta-analisis en Indus Journal of Bioscience Research (2025)'
  },
  'Magnesio': {
    porque: {
      'energia': 'Participa en la produccion de energia y reduce los calambres y la tension muscular. Te vas a sentir menos cansado/a y con mejor recuperacion.',
      'estres': 'Relaja el sistema nervioso y mejora la calidad del sueno. Si te cuesta desconectar o dormis mal, el magnesio es clave.',
      'fitness': 'Mejora la funcion muscular, reduce calambres y acelera la recuperacion post-entreno. Esencial si entrenas con frecuencia.',
      'default': 'Relaja los musculos, mejora el sueno y ayuda a mas de 300 procesos en tu cuerpo. Es uno de los minerales que mas impacto vas a notar.'
    },
    tiempo: '2-4 semanas para sueno, 4-6 semanas para estres',
    evidencia: 'Fuerte',
    papers: 'Meta-analisis en BMC Complementary Medicine (2021)'
  },
  'Magnesio Citrato': {
    porque: {
      'energia': 'Participa en la produccion de energia y reduce los calambres y la tension muscular. Te vas a sentir menos cansado/a y con mejor recuperacion.',
      'estres': 'Relaja el sistema nervioso y mejora la calidad del sueno. Si te cuesta desconectar o dormis mal, el magnesio es clave.',
      'fitness': 'Mejora la funcion muscular, reduce calambres y acelera la recuperacion post-entreno. Esencial si entrenas con frecuencia.',
      'default': 'Relaja los musculos, mejora el sueno y ayuda a mas de 300 procesos en tu cuerpo. Es uno de los minerales que mas impacto vas a notar.'
    },
    tiempo: '2-4 semanas para sueno, 4-6 semanas para estres',
    evidencia: 'Fuerte',
    papers: 'Meta-analisis en BMC Complementary Medicine (2021)'
  },
  'Maca': {
    porque: {
      'energia': 'Es un energizante natural que mejora tu resistencia fisica y mental. Vas a notar mas energia sostenida durante el dia, sin el pico y caida del cafe.',
      'default': 'Mejora tu energia y resistencia de forma natural. Es un adaptogeno: ayuda a tu cuerpo a manejar mejor el estres y la fatiga.'
    },
    tiempo: '4-8 semanas',
    evidencia: 'Moderado',
    papers: 'Revision sistematica en BMC Complementary and Alternative Medicine'
  },
  'Ashwagandha': {
    porque: {
      'estres': 'Reduce el cortisol (la hormona del estres) hasta un 30%. Si te sentis ansioso/a, con la mente acelerada o te cuesta relajarte, es una de las mejores opciones naturales.',
      'energia': 'Ademas de reducir el estres, mejora tu energia porque cuando tu cuerpo no esta en modo "alerta" todo el tiempo, rinde mucho mas.',
      'default': 'Reduce el estres y mejora tu descanso. Vas a notar que te sentis mas tranquilo/a y con mejor calidad de sueno.'
    },
    tiempo: '4-8 semanas',
    evidencia: 'Fuerte',
    papers: 'Meta-analisis en Journal of Clinical Medicine (2022)'
  },
  'Omega 3': {
    porque: {
      'estres': 'Reduce la inflamacion que el estres cronico genera en tu cuerpo. Ademas mejora la concentracion y claridad mental.',
      'piel': 'Hidrata tu piel desde adentro y reduce la inflamacion que causa enrojecimiento o irritacion. Es la base para una piel luminosa.',
      'default': 'Reduce la inflamacion general de tu cuerpo, mejora la concentracion y apoya la salud cardiovascular. Es uno de los suplementos con mas evidencia.'
    },
    tiempo: '4-8 semanas',
    evidencia: 'Moderado-Fuerte',
    papers: 'Meta-analisis en Journal of Clinical Lipidology (2024)'
  },
  'Zinc + Vitamina C': {
    porque: {
      'inmunidad': 'Son la primera linea de defensa de tu sistema inmune. El zinc fortalece tus celulas de defensa y la vitamina C las potencia. Si te enfermades seguido, vas a notar la diferencia.',
      'default': 'Fortalecen tu sistema inmune y te ayudan a enfermarte menos. Es la combinacion clasica de proteccion, pero en la dosis correcta.'
    },
    tiempo: '1-2 semanas para defensa activa',
    evidencia: 'Fuerte',
    papers: 'Meta-analisis Cochrane sobre zinc y resfrios'
  },
  'Probioticos': {
    porque: {
      'inmunidad': 'El 70% de tu sistema inmune esta en tu intestino. Los probioticos equilibran tu flora intestinal para que tus defensas funcionen al maximo.',
      'default': 'Mejoran tu digestion y fortalecen tu sistema inmune. Si tenes hinchacion, gases o digestion lenta, vas a notar un cambio importante.'
    },
    tiempo: '2-4 semanas para digestion, 4-8 para inmunidad',
    evidencia: 'Moderado-Fuerte',
    papers: 'Meta-analisis en Nutrients (2023)'
  },
  'Vitamina D + Calcio': {
    porque: {
      'inmunidad': 'La vitamina D activa tus celulas de defensa. La mayoria de las personas tienen niveles bajos, especialmente en invierno. Es esencial para que tu sistema inmune funcione bien.',
      'default': 'Fortalece tus huesos y tu sistema inmune. La vitamina D es una de las deficiencias mas comunes y corregirla tiene un impacto enorme en como te sentis.'
    },
    tiempo: '4-8 semanas para niveles optimos',
    evidencia: 'Fuerte',
    papers: 'Meta-analisis en Journal of Bone and Mineral Research'
  },
  'Colageno Hidrolizado': {
    porque: {
      'piel': 'Estimula a tu piel a producir mas colageno propio. Vas a notar tu piel mas firme, hidratada y con mejor textura. Tambien fortalece el pelo y las unas.',
      'default': 'Mejora la firmeza y elasticidad de tu piel, fortalece pelo y unas. Es el suplemento estrella para verse y sentirse bien por fuera.'
    },
    tiempo: '4-8 semanas para piel, 8-12 para articulaciones',
    evidencia: 'Moderado',
    papers: 'Meta-analisis en International Journal of Dermatology (2021)'
  },
  'Astaxantina': {
    porque: {
      'piel': 'Es un antioxidante ultra potente que protege tu piel del daño solar y del envejecimiento. Reduce arrugas finas y mejora la luminosidad.',
      'default': 'Protege tus celulas del daño oxidativo. Es uno de los antioxidantes mas potentes que existen, ideal para mantener tu piel joven y tu cuerpo protegido.'
    },
    tiempo: '4-8 semanas',
    evidencia: 'Moderado',
    papers: 'Meta-analisis en Pharmacological Research (2022)'
  },
  'Proteina en Polvo': {
    porque: {
      'fitness': 'Aporta los aminoacidos que tus musculos necesitan para recuperarse y crecer despues del entreno. Sin proteina suficiente, el esfuerzo fisico no se traduce en resultados.',
      'default': 'Ayuda a tu cuerpo a recuperarse y mantener masa muscular. Ideal si no llegas a consumir suficiente proteina en tu dieta diaria.'
    },
    tiempo: '2-4 semanas con entrenamiento consistente',
    evidencia: 'Muy Fuerte',
    papers: 'Meta-analisis en British Journal of Sports Medicine (2018)'
  },
  'Creatina Monohidrato': {
    porque: {
      'fitness': 'Te da mas fuerza y potencia en cada serie. Es el suplemento deportivo con mas evidencia cientifica del mundo. Tambien mejora la concentracion y funcion cognitiva.',
      'default': 'Mejora tu rendimiento fisico y tu funcion cognitiva. Es el suplemento mas estudiado y seguro que existe.'
    },
    tiempo: '2-4 semanas para notar diferencia',
    evidencia: 'Muy Fuerte',
    papers: 'Meta-analisis en Journal of the International Society of Sports Nutrition (2017)'
  }
};

const OBJETIVOS = {
  'energia': 'Energía y Vitalidad',
  'estres': 'Estrés y Sueño',
  'inmunidad': 'Inmunidad',
  'piel': 'Piel y Pelo',
  'fitness': 'Fitness y Rendimiento'
};

// Helper: buscar en un mapeo tolerando tildes y variantes
// Ej: buscarEnMapeo(MARCAS, 'Colageno Hidrolizado') encuentra 'Colágeno Hidrolizado'
function buscarEnMapeo(mapeo, nombre) {
  if (!nombre) return '';
  // Intento directo
  if (mapeo[nombre] !== undefined) return mapeo[nombre];
  // Intento sin tildes
  var sinTildes = nombre.normalize('NFD').replace(/[\u0300-\u036f]/g, '');
  for (var key in mapeo) {
    var keySinTildes = key.normalize('NFD').replace(/[\u0300-\u036f]/g, '');
    if (keySinTildes === sinTildes || keySinTildes === nombre || key === sinTildes) {
      return mapeo[key];
    }
  }
  return '';
}

// Helper: parse objetivo string (puede ser "energia", "energia, estres", o nombre largo)
// Devuelve array de keys: ['energia', 'estres']
function parseObjetivoKeys(objetivoRaw) {
  if (!objetivoRaw) return [];
  var parts = String(objetivoRaw).split(',').map(function(s) { return s.trim(); }).filter(Boolean);
  var keys = [];
  parts.forEach(function(part) {
    var k = OBJETIVO_A_KEY[part] || part.toLowerCase();
    if (OBJETIVOS[k]) keys.push(k);
  });
  // Si no matcheo ninguno, intentar con el string completo
  if (keys.length === 0) {
    var lower = objetivoRaw.toLowerCase();
    if (lower.includes('energ')) keys.push('energia');
    if (lower.includes('estr') || lower.includes('sue')) keys.push('estres');
    if (lower.includes('inmun')) keys.push('inmunidad');
    if (lower.includes('piel') || lower.includes('pelo')) keys.push('piel');
    if (lower.includes('fitness') || lower.includes('rend')) keys.push('fitness');
  }
  return keys.length > 0 ? keys : ['default'];
}

// Helper: nombre legible para objetivos multiples
// "energia, estres" -> "Energía y Vitalidad y Estrés y Sueño"
function objetivoNombreDisplay(objetivoRaw) {
  var keys = parseObjetivoKeys(objetivoRaw);
  var nombres = keys.filter(function(k) { return k !== 'default'; }).map(function(k) { return OBJETIVOS[k] || k; });
  if (nombres.length === 0) return 'bienestar';
  if (nombres.length === 1) return nombres[0];
  return nombres.slice(0, -1).join(', ') + ' y ' + nombres[nombres.length - 1];
}

// Helper: mejor texto "porque" para un suplemento dado multiples objetivos
// Prioriza el primer objetivo que tenga texto especifico, sino default
function mejorPorque(cienciaItem, objetivoKeys) {
  if (!cienciaItem || !cienciaItem.porque) return '';
  for (var i = 0; i < objetivoKeys.length; i++) {
    if (cienciaItem.porque[objetivoKeys[i]]) return cienciaItem.porque[objetivoKeys[i]];
  }
  return cienciaItem.porque['default'] || '';
}


// Helper: envia email desde hola@niva.uy usando GmailApp + "Enviar como"
function nivaSend(to, subject, body, htmlBody) {
  if (!to || !to.includes('@')) {
    Logger.log('nivaSend: email destinatario vacio o invalido: ' + to);
    return;
  }
  GmailApp.sendEmail(to, subject, body || '', {
    htmlBody: htmlBody || undefined,
    from: NIVA_EMAIL,
    name: 'niva - tu bienestar',
    replyTo: NIVA_REPLY_TO
  });
}


// ===== FUNCIÓN PRINCIPAL =====
function onFormSubmit(e) {
  try {
    const responses = e.response.getItemResponses();

    // Build a map by index (position) since titles may vary
    const byIndex = responses.map(r => r.getResponse());

    // Also build by title for flexibility
    const byTitle = {};
    responses.forEach(r => {
      byTitle[r.getItem().getTitle().trim()] = r.getResponse();
    });

    // Log all field titles for debugging
    Logger.log('Form field titles: ' + responses.map(r => r.getItem().getTitle()).join(' | '));
    Logger.log('Form values: ' + byIndex.join(' | '));

    // Parse fields — try by title first, then by position
    // Expected field order: Nombre(0), Email(1), Objetivo(2), Sup1(3), Dosis1(4),
    //   Sup2(5), Dosis2(6), Sup3(7), Dosis3(8), TipoCompra(9), Precio(10), JSON(11)
    const nombre = byTitle['Nombre'] || byTitle['nombre'] || byIndex[0] || '';
    const emailRaw = byTitle['Email'] || byTitle['email'] || byIndex[1] || '';
    const email = emailRaw.replace(/\s+/g, '').toLowerCase(); // limpiar espacios y normalizar
    const objetivoRaw = byTitle['Objetivo'] || byTitle['objetivo'] || byIndex[2] || '';
    const sup1 = byTitle['Sup1'] || byIndex[3] || '';
    const dosis1 = byTitle['Dosis1'] || byIndex[4] || '';
    const sup2 = byTitle['Sup2'] || byIndex[5] || '';
    const dosis2 = byTitle['Dosis2'] || byIndex[6] || '';
    const sup3 = byTitle['Sup3'] || byIndex[7] || '';
    const dosis3 = byTitle['Dosis3'] || byIndex[8] || '';
    const tipoCompra = byTitle['Tipo Compra'] || byTitle['tipo_compra'] || byIndex[9] || '';
    const precioRaw = byTitle['Precio'] || byTitle['precio'] || byIndex[10] || '';
    const jsonData = byTitle['JSON'] || byTitle['json'] || byIndex[11] || '{}';

    // Parse JSON blob (contains all structured data as backup)
    let parsed = {};
    try { parsed = JSON.parse(jsonData); } catch(ex) {
      Logger.log('Error parsing JSON: ' + ex.toString());
    }

    const firstName = nombre.split(' ')[0] || 'Cliente';

    // Objetivo: try raw value, then from JSON — puede ser multi-objetivo "energia, estres"
    const objetivoKey = objetivoRaw || parsed.objetivo || '';
    const objetivoNombre = objetivoNombreDisplay(objetivoKey);

    // Dirección: from JSON
    const direccion = parsed.direccion || 'Pendiente de confirmar';

    // Frecuencia y plan: from JSON first, then from individual fields
    const frecuencia = parsed.frecuencia || 'mensual';
    const planNombre = parsed.plan || 'Plan Esencial';

    // Precio: from JSON first (number), then from individual field
    let precio = '';
    if (parsed.precio && parsed.precio > 0) {
      precio = String(parsed.precio);
    } else if (precioRaw && precioRaw !== '0' && precioRaw !== '') {
      precio = precioRaw;
    }
    // Format precio with dot separator for thousands
    if (precio) {
      const num = parseInt(precio);
      if (!isNaN(num)) {
        precio = num.toLocaleString('es-UY');
      }
    }

    // Build supplements list — from individual fields first
    const sups = [];
    if (sup1) sups.push({ nombre: sup1, dosis: dosis1 });
    if (sup2) sups.push({ nombre: sup2, dosis: dosis2 });
    if (sup3) sups.push({ nombre: sup3, dosis: dosis3 });

    // If no individual sups found, try parsing from JSON suplementos array
    if (sups.length === 0 && parsed.suplementos && parsed.suplementos.length > 0) {
      parsed.suplementos.forEach(function(item) {
        const match = item.match(/^(.+?)\s*\((.+)\)$/);
        if (match) {
          sups.push({ nombre: match[1].trim(), dosis: match[2].trim() });
        } else {
          sups.push({ nombre: item.trim(), dosis: '' });
        }
      });
    } else if (parsed.suplementos && parsed.suplementos.length > 3) {
      // Add extras from JSON beyond the first 3
      for (let i = 3; i < parsed.suplementos.length; i++) {
        const match = parsed.suplementos[i].match(/^(.+?)\s*\((.+)\)$/);
        if (match) sups.push({ nombre: match[1].trim(), dosis: match[2].trim() });
        else sups.push({ nombre: parsed.suplementos[i].trim(), dosis: '' });
      }
    }

    // Detectar tipo de envío:
    // - Lead puro: vio resultados pero no clickeó comprar (no tiene precio ni dirección)
    // - Pendiente de pago: clickeó comprar, fue a MercadoPago (tiene precio y dirección pero tipo='pendiente')
    var hasPrecio = precio && precio !== '0' && precio !== '';
    var hasDireccion = direccion && direccion !== 'Pendiente de confirmar';
    var isPendientePago = (tipoCompra === 'pendiente') && (hasPrecio || hasDireccion);
    var isLeadPuro = !isPendientePago;

    Logger.log('Parsed data — nombre: ' + nombre + ' | email: ' + email +
      ' | objetivo: ' + objetivoNombre + ' | plan: ' + planNombre +
      ' | frecuencia: ' + frecuencia + ' | precio: ' + precio +
      ' | tipo: ' + (isPendientePago ? 'PENDIENTE_PAGO' : 'LEAD') +
      ' | sups: ' + sups.map(function(s) { return s.nombre; }).join(', '));

    if (isPendientePago) {
      // ===== PENDIENTE DE PAGO — clickeó comprar, fue a MercadoPago =====
      sendPendingEmail(email, firstName, planNombre, frecuencia, precio, sups, objetivoNombre);
      // Notificación interna con asunto claro de "pendiente de pago"
      var pendBody = 'SOLICITUD PENDIENTE DE PAGO\n\n' +
        'Nombre: ' + nombre + '\n' +
        'Email: ' + email + '\n' +
        'Plan: ' + planNombre + ' (' + frecuencia + ')\n' +
        'Precio: $' + precio + '/mes\n\n' +
        'Suplementos:\n' + sups.map(function(s) { return '  - ' + s.nombre + ': ' + (s.dosis || ''); }).join('\n') + '\n\n' +
        'Dirección: ' + direccion + '\n\n' +
        '---\nAcción: verificar pago en MercadoPago y cambiar a "confirmado" en el spreadsheet.';
      GmailApp.sendEmail(NIVA_ADMIN, 'Pendiente de pago: ' + nombre + ' - ' + planNombre, pendBody, {
        from: NIVA_EMAIL, name: 'niva - pedidos', replyTo: NIVA_EMAIL
      });

    } else {
      // ===== LEAD PURO — vio resultados pero no compró =====
      Logger.log('>>> LEAD DETECTADO — enviando sendLeadEmail a: ' + email + ' con ' + sups.length + ' sups');
      sendLeadEmail(email, firstName, objetivoNombre, sups);
      Logger.log('Lead registrado + email de recomendación enviado OK: ' + email);
    }

  } catch (error) {
    Logger.log('Error en onFormSubmit: ' + error.toString());
    // Enviar error a admin directo (evita deduplicación Cloudflare)
    GmailApp.sendEmail(NIVA_ADMIN, 'ERROR en pedido niva', 'Error procesando un pedido:\n\n' + error.toString() + '\n\nStack: ' + (error.stack || 'N/A'), {
      from: NIVA_EMAIL, name: 'niva - errores'
    });
  }
}


// ===== EMAIL LEAD: Recomendación + precio + botón de compra directo =====
function sendLeadEmail(email, nombre, objetivo, sups) {
  // Mapear objetivo(s) a keys — soporta multi-objetivo
  var objKeys = parseObjetivoKeys(objetivo);

  var supsHtml = '';
  var supIds = []; // para el deep link
  sups.forEach(function(s) {
    var color = buscarEnMapeo(DISC_COLORS, s.nombre) || '#B8C1B6';
    var label = buscarEnMapeo(DISC_LABELS, s.nombre) || '•';
    var ciencia = buscarEnMapeo(CIENCIA, s.nombre) || {};
    var supId = buscarEnMapeo(NOMBRE_A_ID, s.nombre) || '';
    if (supId) supIds.push(supId);

    // Texto personalizado segun objetivo(s) — usa el mejor match
    var porQue = mejorPorque(ciencia, objKeys);
    // Acortar a la primera oracion para el lead
    if (porQue.indexOf('.') > 0) porQue = porQue.substring(0, porQue.indexOf('.') + 1);

    supsHtml += '<div style="margin-bottom:14px;">' +
      '<table cellpadding="0" cellspacing="0"><tr>' +
      '<td style="width:40px; height:40px; border-radius:50%; background:' + color + '; text-align:center; vertical-align:middle; font-weight:600; font-size:14px; color:rgba(46,46,46,0.55); line-height:40px;">' + label + '</td>' +
      '<td style="padding-left:12px; vertical-align:middle;">' +
      '<div style="font-size:15px; font-weight:500; color:#2E2E2E;">' + s.nombre + ' <span style="font-weight:400; font-size:12px; color:#7A7972;">· ' + (buscarEnMapeo(MARCAS, s.nombre) || '') + '</span></div>' +
      '<div style="font-size:11px; color:#A8A39A; font-style:italic; margin-top:1px;">' + (buscarEnMapeo(QUE_ES, s.nombre) || '') + '</div>' +
      '<div style="font-size:11px; color:#A8A39A; margin-top:1px;">' + (buscarEnMapeo(FORMATOS, s.nombre) || '') + '</div>' +
      '<div style="font-size:13px; color:#7A7972; line-height:1.5; margin-top:2px;">' + porQue + '</div>' +
      '</td></tr></table></div>';
  });

  // Armar deep link: niva.uy?n=nombre&e=email&o=key1,key2&s=id1,id2,id3
  var objetivoKeyStr = objKeys.filter(function(k) { return k !== 'default'; }).join(',') || 'energia';
  var baseSupIds = supIds.slice(0, 3).join(',');
  var extraSupIds = supIds.slice(3).join(',');
  var deepLink = 'https://niva.uy?n=' + encodeURIComponent(nombre) +
    '&e=' + encodeURIComponent(email) +
    '&o=' + encodeURIComponent(objetivoKeyStr) +
    '&s=' + encodeURIComponent(baseSupIds);
  if (extraSupIds) deepLink += '&x=' + encodeURIComponent(extraSupIds);

  // Determinar plan y precio según cantidad de sups
  var numBase = 3;
  var numExtras = Math.max(sups.length - 3, 0);
  var planNombre = numExtras > 0 ? 'Plan Completo' : 'Plan Esencial';
  var planPrecio = numExtras > 0 ? 5990 : 3990;
  var planPrecioStr = planPrecio.toLocaleString('es-UY');
  var planDetalle = numExtras > 0 ? (numBase + ' base + ' + numExtras + ' adicional' + (numExtras > 1 ? 'es' : '')) : (numBase + ' suplementos base');

  var html = '<div style="font-family:\'Helvetica Neue\',Arial,sans-serif; max-width:560px; margin:0 auto; background:#fff; border-radius:16px; overflow:hidden;">' +
    '<div style="padding:32px 40px 20px; text-align:center;">' +
    '<div style="font-size:17px; font-weight:600; color:#2E2E2E; letter-spacing:4px;">niva</div>' +
    '</div>' +
    // Saludo + intro
    '<div style="padding:0 40px 24px;">' +
    '<h1 style="font-size:26px; font-weight:300; color:#2E2E2E; margin:0 0 8px;">Tu recomendación personalizada</h1>' +
    '<p style="font-size:15px; color:#4A4A4A; line-height:1.6;">' + nombre + ', analizamos tu perfil y seleccionamos estos suplementos para tu objetivo de <strong>' + objetivo.toLowerCase() + '</strong>:</p>' +
    '</div>' +
    // Lista de suplementos con por qué
    '<div style="padding:0 40px 24px;">' +
    '<div style="background:#F6F4F2; border-radius:12px; padding:20px;">' +
    '<div style="font-size:13px; font-weight:600; color:#6E766B; letter-spacing:1px; text-transform:uppercase; margin-bottom:14px;">Tu selección personalizada</div>' +
    supsHtml +
    '</div></div>' +
    // Bloque de precio + plan
    '<div style="padding:0 40px 24px;">' +
    '<div style="background:#fff; border:1.5px solid #E9E6E1; border-radius:12px; padding:20px; text-align:center;">' +
    '<div style="font-size:13px; font-weight:600; color:#6E766B; letter-spacing:1px; text-transform:uppercase; margin-bottom:8px;">' + planNombre + '</div>' +
    '<div style="font-size:14px; color:#7A7972; margin-bottom:8px;">' + planDetalle + ' · envío gratis</div>' +
    '<div style="font-size:28px; font-weight:500; color:#2E2E2E;">$' + planPrecioStr + ' <span style="font-size:14px; font-weight:400; color:#7A7972;">UYU/mes</span></div>' +
    '<div style="font-size:12px; color:#9A9A92; margin-top:4px;">Incluye dosis personalizadas, horarios de toma y seguimiento semanal</div>' +
    '</div></div>' +
    // Botón de compra
    '<div style="padding:0 40px 12px; text-align:center;">' +
    '<a href="' + deepLink + '" style="display:inline-block; padding:16px 48px; background:#2E2E2E; color:#fff; border-radius:24px; text-decoration:none; font-size:16px; font-weight:500;">Activar mi plan</a>' +
    '</div>' +
    '<div style="padding:0 40px 24px; text-align:center;">' +
    '<p style="font-size:12px; color:#9A9A92;">Al hacer click vas a ver tu plan completo y podés confirmar con pago seguro por Mercado Pago</p>' +
    '</div>' +
    // Dudas
    '<div style="padding:0 40px 24px; text-align:center;">' +
    '<p style="font-size:13px; color:#7A7972;">¿Tenés dudas? Escribinos a <a href="mailto:hola@niva.uy" style="color:#6E766B; font-weight:500;">hola@niva.uy</a> o por <a href="https://instagram.com/nivauy" style="color:#6E766B; font-weight:500;">@nivauy</a></p>' +
    '</div>' +
    // Footer
    '<div style="padding:24px 40px; background:#2E2E2E; text-align:center;">' +
    '<div style="font-size:14px; font-weight:600; color:#fff; letter-spacing:3px;">niva</div>' +
    '<div style="font-size:12px; color:rgba(255,255,255,0.6);">tu bienestar, más simple</div>' +
    '<div style="font-size:11px; color:rgba(255,255,255,0.5); margin-top:8px;">Recomendaciones elaboradas por Lic. en Nutrición</div>' +
    '<div style="font-size:12px; color:rgba(255,255,255,0.4); margin-top:8px;">hola@niva.uy · niva.uy · <a href="https://instagram.com/nivauy" style="color:rgba(255,255,255,0.4);">@nivauy</a></div>' +
    '</div></div>';

  nivaSend(email, nombre + ', tu recomendación niva está lista', '', html);
}


// ===== EMAIL 0: Solicitud recibida (pendiente de pago) =====
function sendPendingEmail(email, nombre, plan, frecuencia, precio, sups, objetivo) {
  let supsHtml = '';
  var supIds = [];
  sups.forEach(function(s) {
    var color = buscarEnMapeo(DISC_COLORS, s.nombre) || '#B8C1B6';
    var label = buscarEnMapeo(DISC_LABELS, s.nombre) || '•';
    var supId = buscarEnMapeo(NOMBRE_A_ID, s.nombre) || '';
    if (supId) supIds.push(supId);
    supsHtml += '<div style="margin-bottom:8px;">' +
      '<table cellpadding="0" cellspacing="0"><tr>' +
      '<td style="width:36px; height:36px; border-radius:50%; background:' + color + '; text-align:center; vertical-align:middle; font-weight:600; font-size:13px; color:rgba(46,46,46,0.55); line-height:36px;">' + label + '</td>' +
      '<td style="padding-left:10px; vertical-align:middle;">' +
      '<div style="font-size:14px; font-weight:500; color:#2E2E2E;">' + s.nombre + ' <span style="font-weight:400; font-size:11px; color:#7A7972;">· ' + (buscarEnMapeo(MARCAS, s.nombre) || '') + '</span></div>' +
      '<div style="font-size:11px; color:#A8A39A; font-style:italic;">' + (buscarEnMapeo(QUE_ES, s.nombre) || '') + '</div>' +
      '<div style="font-size:11px; color:#A8A39A;">' + (buscarEnMapeo(FORMATOS, s.nombre) || '') + '</div>' +
      '</td></tr></table></div>';
  });

  // Armar deep link — soporta multi-objetivo
  var objKeys = parseObjetivoKeys(objetivo);
  var objetivoKeyStr = objKeys.filter(function(k) { return k !== 'default'; }).join(',') || 'energia';
  var baseSupIds = supIds.slice(0, 3).join(',');
  var extraSupIds = supIds.slice(3).join(',');
  var deepLink = 'https://niva.uy?n=' + encodeURIComponent(nombre) +
    '&e=' + encodeURIComponent(email) +
    '&o=' + encodeURIComponent(objetivoKeyStr) +
    '&s=' + encodeURIComponent(baseSupIds);
  if (extraSupIds) deepLink += '&x=' + encodeURIComponent(extraSupIds);

  const frecMap = { mensual: 'Mensual', trimestral: 'Trimestral (-10%)', semestral: 'Semestral (-15%)' };
  const precioDisplay = precio ? ('$' + precio + ' UYU/mes') : '';

  const html = '<div style="font-family:\'Helvetica Neue\',Arial,sans-serif; max-width:560px; margin:0 auto; background:#fff; border-radius:16px; overflow:hidden;">' +
    '<div style="padding:32px 40px 20px; text-align:center;">' +
    '<div style="font-size:17px; font-weight:600; color:#2E2E2E; letter-spacing:4px;">niva</div>' +
    '</div>' +
    '<div style="padding:0 40px 24px;">' +
    '<h1 style="font-size:26px; font-weight:300; color:#2E2E2E; margin:0 0 8px;">Hola ' + nombre + ',</h1>' +
    '<p style="font-size:15px; color:#4A4A4A; line-height:1.6;">Recibimos tu solicitud. Una vez que se acredite tu pago, te enviamos la confirmacion y tu plan personalizado con dosis, horarios y la ciencia detras de cada recomendacion.</p>' +
    '</div>' +
    '<div style="padding:0 40px 24px;">' +
    '<div style="background:#F6F4F2; border-radius:12px; padding:20px;">' +
    '<div style="font-size:13px; font-weight:600; color:#6E766B; letter-spacing:1px; text-transform:uppercase; margin-bottom:12px;">Tu seleccion</div>' +
    '<div style="font-size:14px; color:#2E2E2E; font-weight:500;">' + plan + ' - ' + (frecMap[frecuencia] || frecuencia) + '</div>' +
    (precioDisplay ? '<div style="font-size:20px; font-weight:500; color:#2E2E2E; margin:4px 0 16px;">' + precioDisplay + '</div>' : '') +
    '<div style="border-top:1px solid #E9E6E1; padding-top:16px;">' +
    supsHtml +
    '</div></div></div>' +
    // Boton de activar plan / deep link
    '<div style="padding:0 40px 24px; text-align:center;">' +
    '<p style="font-size:14px; color:#4A4A4A; line-height:1.6; margin-bottom:16px;">Si ya completaste el pago, no necesitas hacer nada mas. Te confirmamos en breve.</p>' +
    '<p style="font-size:14px; color:#4A4A4A; line-height:1.6; margin-bottom:16px;">Si tuviste algun problema o queres retomar tu compra:</p>' +
    '<a href="' + deepLink + '" style="display:inline-block; background:#2E2E2E; color:#fff; padding:14px 32px; border-radius:10px; text-decoration:none; font-size:15px; font-weight:500; letter-spacing:0.5px;">Activar mi plan</a>' +
    '</div>' +
    '<div style="padding:0 40px 24px; text-align:center;">' +
    '<p style="font-size:13px; color:#7A7972;">O escribinos a <a href="mailto:hola@niva.uy" style="color:#6E766B; font-weight:500;">hola@niva.uy</a></p>' +
    '</div>' +
    '<div style="padding:24px 40px; background:#2E2E2E; text-align:center;">' +
    '<div style="font-size:14px; font-weight:600; color:#fff; letter-spacing:3px;">niva</div>' +
    '<div style="font-size:12px; color:rgba(255,255,255,0.6);">tu bienestar, mas simple</div>' +
    '<div style="font-size:12px; color:rgba(255,255,255,0.4); margin-top:12px;">hola@niva.uy - niva.uy - <a href="https://instagram.com/nivauy" style="color:rgba(255,255,255,0.4);">@nivauy</a></div>' +
    '</div></div>';

  nivaSend(email, nombre + ', recibimos tu solicitud - niva', '', html);
}


// ===== EMAIL 1: Confirmación (se envía solo cuando se confirma el pago) =====
// sendConfirmationEmail ya no se usa por separado — se fusiono con sendPlanEmail
// Se mantiene como wrapper por compatibilidad pero no envia nada
function sendConfirmationEmail(email, nombre, plan, frecuencia, precio, sups, direccion) {
  // No hace nada — la confirmacion ahora va dentro de sendPlanEmail
  Logger.log('sendConfirmationEmail: saltado (fusionado con sendPlanEmail)');
}


// ===== EMAIL 2: Confirmacion + Plan personalizado (fusionados) =====
function sendPlanEmail(email, nombre, objetivo, sups, planNombre, frecuencia, precio, direccion) {
  // Mapear objetivo(s) a keys — soporta multi-objetivo
  var objKeys = parseObjetivoKeys(objetivo);

  let supsHtml = '';
  sups.forEach(function(s) {
    var color = buscarEnMapeo(DISC_COLORS, s.nombre) || '#B8C1B6';
    var label = buscarEnMapeo(DISC_LABELS, s.nombre) || '•';
    var forma = FORMAS[s.nombre] || 'capsula';
    var ciencia = buscarEnMapeo(CIENCIA, s.nombre) || {};

    // Texto personalizado segun objetivo(s) — usa el mejor match
    var porqueTexto = mejorPorque(ciencia, objKeys);

    supsHtml += '<div style="padding:0 40px 16px;">' +
      '<div style="border:1px solid #E9E6E1; border-radius:12px; padding:20px;">' +
      '<div style="margin-bottom:12px;">' +
      '<table cellpadding="0" cellspacing="0"><tr>' +
      '<td style="width:48px; height:48px; border-radius:50%; background:' + color + '; text-align:center; vertical-align:middle; font-weight:600; font-size:16px; color:rgba(46,46,46,0.55); line-height:48px;">' + label + '</td>' +
      '<td style="padding-left:14px; vertical-align:middle;">' +
      '<div style="font-size:16px; font-weight:500; color:#2E2E2E;">' + s.nombre + ' <span style="font-weight:400; font-size:12px; color:#7A7972;">· ' + (buscarEnMapeo(MARCAS, s.nombre) || '') + '</span></div>' +
      '<div style="font-size:12px; color:#7A7972; font-style:italic;">' + (buscarEnMapeo(QUE_ES, s.nombre) || '') + '</div>' +
      '<div style="font-size:12px; color:#7A7972;">' + (buscarEnMapeo(FORMATOS, s.nombre) || forma) + '</div>' +
      '</td></tr></table></div>' +
      '<table width="100%" cellpadding="0" cellspacing="0" style="margin-bottom:12px;"><tr>' +
      '<td style="padding:8px 12px; background:#F6F4F2; border-radius:8px; width:48%;">' +
      '<div style="font-size:11px; color:#7A7972; text-transform:uppercase;">Dosis</div>' +
      '<div style="font-size:13px; color:#2E2E2E; font-weight:500; margin-top:2px;">' + (s.dosis || 'Segun indicacion') + '</div></td>' +
      '<td style="width:4%;"></td>' +
      '<td style="padding:8px 12px; background:#F6F4F2; border-radius:8px; width:48%;">' +
      '<div style="font-size:11px; color:#7A7972; text-transform:uppercase;">Cuando tomar</div>' +
      '<div style="font-size:13px; color:#2E2E2E; font-weight:500; margin-top:2px;">' + (buscarEnMapeo(CUANDO_TOMAR, s.nombre) || 'Con una comida principal') + '</div></td>' +
      '</tr></table>' +
      '<div style="font-size:13px; color:#4A4A4A; line-height:1.6; margin-bottom:8px;">' +
      '<strong style="color:#6E766B;">Por que te lo recomendamos:</strong> ' + porqueTexto + '</div>' +
      '<div style="font-size:12px; color:#4A4A4A; line-height:1.6; margin-bottom:8px;">' +
      '<strong style="color:#6E766B;">Cuando notas el efecto:</strong> ' + (ciencia.tiempo || '') + '</div>' +
      '<div style="font-size:11px; color:#7A7972; line-height:1.5; padding-top:8px; border-top:1px solid #E9E6E1;">' +
      '<strong>Respaldo cientifico (' + (ciencia.evidencia || '') + '):</strong> ' + (ciencia.papers || '') + '</div>' +
      '</div></div>';
  });

  // Fallback if no supplements
  if (sups.length === 0) {
    supsHtml = '<div style="padding:0 40px 16px;"><div style="font-size:14px; color:#7A7972;">Estamos preparando tu rutina personalizada. Te la enviamos pronto a este email.</div></div>';
  }

  // Build intro text — handle blank/multi objetivo gracefully
  var objDisplay = objetivoNombreDisplay ? objetivoNombreDisplay(objetivo) : objetivo;
  var objPalabra = (objetivo && objetivo.includes(',')) ? 'tus objetivos de' : 'tu objetivo de';
  let introText = '';
  if (objetivo && objetivo !== 'bienestar' && objetivo !== '') {
    introText = nombre + ', acá tenés tu guía completa. Cada suplemento fue seleccionado por nuestra Lic. en Nutrición para ' + objPalabra + ' <strong>' + objDisplay.toLowerCase() + '</strong>, con dosis y horarios respaldados por evidencia científica. Todos los productos son de marcas reconocidas, llegan en su envase original sellado, y las dosis están dentro de los rangos seguros.';
  } else {
    introText = nombre + ', acá tenés tu guía completa. Cada suplemento fue seleccionado por nuestra Lic. en Nutrición para vos, con dosis y horarios respaldados por evidencia científica. Todos los productos son de marcas reconocidas, llegan en su envase original sellado, y las dosis están dentro de los rangos seguros.';
  }

  // Seccion de confirmacion (solo si viene con datos de compra)
  var frecMap = { mensual: 'Mensual', trimestral: 'Trimestral (-10%)', semestral: 'Semestral (-15%)' };
  var confirmBlock = '';
  if (planNombre && precio) {
    var precioDisplay = '$' + precio + ' UYU/mes';
    confirmBlock =
    '<div style="padding:0 40px 24px;">' +
    '<div style="background:#f0f7f0; border-radius:12px; padding:20px; border-left:4px solid #8FA08A;">' +
    '<div style="font-size:13px; font-weight:600; color:#6E766B; letter-spacing:1px; text-transform:uppercase; margin-bottom:8px;">Pago confirmado</div>' +
    '<div style="font-size:14px; color:#2E2E2E; font-weight:500;">' + planNombre + ' - ' + (frecMap[frecuencia] || frecuencia) + '</div>' +
    '<div style="font-size:18px; font-weight:500; color:#2E2E2E; margin:4px 0;">' + precioDisplay + '</div>' +
    (direccion ? '<div style="font-size:13px; color:#7A7972; margin-top:8px; padding-top:8px; border-top:1px solid #d4e4d4;">Envio a: ' + direccion + '</div>' : '') +
    '</div></div>';
  }

  const html = '<div style="font-family:\'Helvetica Neue\',Arial,sans-serif; max-width:560px; margin:0 auto; background:#fff; border-radius:16px; overflow:hidden;">' +
    '<div style="padding:32px 40px 20px; text-align:center;">' +
    '<div style="font-size:17px; font-weight:600; color:#2E2E2E; letter-spacing:4px;">niva</div>' +
    '</div>' +
    '<div style="padding:0 40px 24px;">' +
    '<h1 style="font-size:26px; font-weight:300; color:#2E2E2E; margin:0 0 8px;">Tu plan personalizado</h1>' +
    '<p style="font-size:15px; color:#4A4A4A; line-height:1.6;">' + introText + '</p>' +
    '</div>' +
    confirmBlock +
    '<div style="padding:0 40px 8px;">' +
    '<div style="font-size:13px; font-weight:600; color:#6E766B; letter-spacing:1px; text-transform:uppercase;">Tu rutina diaria</div>' +
    '</div>' +
    supsHtml +
    '<div style="padding:16px 40px 24px;">' +
    '<div style="background:#F6F4F2; border-radius:12px; padding:20px;">' +
    '<div style="font-size:13px; font-weight:600; color:#6E766B; letter-spacing:1px; text-transform:uppercase; margin-bottom:10px;">Tips para empezar</div>' +
    '<p style="font-size:14px; color:#4A4A4A; line-height:1.7; margin:0 0 6px;">· Empezá de a uno: incorporá un suplemento nuevo cada 2-3 días.</p>' +
    '<p style="font-size:14px; color:#4A4A4A; line-height:1.7; margin:0 0 6px;">· Consistencia es clave: los efectos se notan entre 2 y 6 semanas.</p>' +
    '<p style="font-size:14px; color:#4A4A4A; line-height:1.7; margin:0 0 6px;">· Tomá con agua abundante y junto con comida (salvo en ayunas).</p>' +
    '<p style="font-size:14px; color:#4A4A4A; line-height:1.7; margin:0;">· Si sentís alguna molestia, escribinos a hola@niva.uy</p>' +
    '</div></div>' +
    '<div style="padding:0 40px 32px;">' +
    '<p style="font-size:14px; color:#4A4A4A; line-height:1.6;">En una semana te enviamos tu primer <strong>check-in</strong>: 5 preguntas rápidas para ir ajustando tu plan según cómo te sentís.</p>' +
    '</div>' +
    '<div style="padding:24px 40px; background:#2E2E2E; text-align:center;">' +
    '<div style="font-size:14px; font-weight:600; color:#fff; letter-spacing:3px;">niva</div>' +
    '<div style="font-size:12px; color:rgba(255,255,255,0.6);">tu bienestar, más simple</div>' +
    '<div style="font-size:12px; color:rgba(255,255,255,0.4); margin-top:12px;">hola@niva.uy · niva.uy · <a href="https://instagram.com/nivauy" style="color:rgba(255,255,255,0.4);">@nivauy</a></div>' +
    '<div style="font-size:11px; color:rgba(255,255,255,0.3); margin-top:12px;">Esta información es educativa. Consultá con tu médico antes de comenzar cualquier suplementación.</div>' +
    '</div></div>';

  var asunto = (planNombre && precio) ?
    nombre + ', tu compra esta confirmada - aca esta tu plan niva' :
    nombre + ', tu plan niva personalizado';
  nivaSend(email, asunto, '', html);

  // Guardar el HTML del plan para usarlo en la notificacion interna
  return html;
}




// ===== NOTIFICACION INTERNA — UN SOLO EMAIL CON TODO =====
function sendInternalNotification(nombre, email, plan, frecuencia, precio, sups, direccion, planHtml) {
  sups = sups || [];

  var supsRows = sups.map(function(s) {
    var cuando = buscarEnMapeo(CUANDO_TOMAR, s.nombre) || '';
    return '<tr>' +
      '<td style="padding:8px 12px; border-bottom:1px solid #eee; font-size:14px;">[ ] ' + s.nombre + '</td>' +
      '<td style="padding:8px 12px; border-bottom:1px solid #eee; font-size:13px; color:#666;">' + (s.dosis || '') + '</td>' +
      '<td style="padding:8px 12px; border-bottom:1px solid #eee; font-size:13px; color:#666;">' + cuando + '</td>' +
      '</tr>';
  }).join('');

  var html = '<div style="font-family:Arial,sans-serif; max-width:600px; margin:0 auto;">' +

    '<div style="background:#2E2E2E; color:#fff; padding:20px; text-align:center;">' +
    '<div style="font-size:20px; font-weight:600; letter-spacing:3px;">niva</div>' +
    '<div style="font-size:14px; color:#ccc; margin-top:4px;">COMPRA CONFIRMADA - ARMAR PAQUETE</div>' +
    '</div>' +

    '<div style="background:#F6F4F2; padding:20px;">' +
    '<h3 style="margin:0 0 12px; font-size:16px; color:#2E2E2E;">DATOS DEL PEDIDO</h3>' +
    '<table style="width:100%; font-size:14px; color:#333;">' +
    '<tr><td style="padding:4px 0; font-weight:600; width:120px;">Cliente:</td><td>' + nombre + '</td></tr>' +
    '<tr><td style="padding:4px 0; font-weight:600;">Email:</td><td>' + email + '</td></tr>' +
    '<tr><td style="padding:4px 0; font-weight:600;">Plan:</td><td>' + plan + ' (' + frecuencia + ')</td></tr>' +
    '<tr><td style="padding:4px 0; font-weight:600;">Precio:</td><td>$' + precio + '/mes</td></tr>' +
    '</table>' +
    '</div>' +

    '<div style="background:#FFFFCC; padding:16px 20px; border-left:4px solid #FFD700;">' +
    '<h3 style="margin:0 0 8px; font-size:16px; color:#333;">DIRECCION DE ENVIO</h3>' +
    '<p style="margin:0; font-size:15px; color:#333; font-weight:600;">' + direccion + '</p>' +
    '</div>' +

    '<div style="padding:20px;">' +
    '<h3 style="margin:0 0 12px; font-size:16px; color:#2E2E2E;">SUPLEMENTOS A INCLUIR EN LA CAJA</h3>' +
    '<table style="width:100%; border-collapse:collapse;">' +
    '<tr style="background:#eee;">' +
    '<th style="padding:8px 12px; text-align:left; font-size:13px;">Suplemento</th>' +
    '<th style="padding:8px 12px; text-align:left; font-size:13px;">Dosis</th>' +
    '<th style="padding:8px 12px; text-align:left; font-size:13px;">Cuando tomar</th>' +
    '</tr>' +
    supsRows +
    '</table>' +
    '<p style="font-size:12px; color:#999; margin-top:12px;">Total: ' + sups.length + ' suplemento' + (sups.length > 1 ? 's' : '') + '</p>' +
    '</div>' +

    '<div style="padding:16px 20px; background:#f0f7f0; border-left:4px solid #8FA08A;">' +
    '<h3 style="margin:0 0 8px; font-size:15px; color:#333;">CHECKLIST DE ARMADO</h3>' +
    '<p style="margin:0; font-size:14px; color:#555; line-height:1.8;">' +
    '[ ] Suplementos verificados y completos<br>' +
    '[ ] Tarjeta del plan impresa (ver abajo)<br>' +
    '[ ] Pastillero incluido (si primer envio o suscripcion)<br>' +
    '[ ] Caja cerrada con sticker niva<br>' +
    '[ ] Envio coordinado<br>' +
    '[ ] Cambiar estado a "enviado" en el spreadsheet' +
    '</p>' +
    '</div>' +

    '<div style="padding:16px 20px; text-align:center; font-size:12px; color:#999;">' +
    'Fecha: ' + new Date().toLocaleDateString('es-UY') + ' - niva.uy' +
    '</div>';

  // Si tenemos el HTML del plan, agregarlo como seccion imprimible
  if (planHtml) {
    html += '<div style="border-top:3px dashed #ccc; margin:20px 0;"></div>' +
      '<div style="padding:12px; background:#FFFFCC; text-align:center; font-family:Arial; font-size:13px; color:#333;">' +
      'TARJETA PARA IMPRIMIR Y METER EN LA CAJA' +
      '</div>' +
      planHtml;
  }

  html += '</div>';

  GmailApp.sendEmail(NIVA_ADMIN, 'COMPRA CONFIRMADA - ' + nombre + ' - ' + plan, '', {
    htmlBody: html,
    from: NIVA_EMAIL,
    name: 'niva - pedidos',
    replyTo: NIVA_EMAIL
  });
}


// ===== CONFIRMAR COMPRA (se ejecuta automáticamente cada 5 min) =====
// Solo tenés que hacer UNA cosa: en el spreadsheet, cambiar "pendiente" por "confirmado".
// Esta función detecta eso y manda los emails automáticamente.
//
// SETUP DEL TRIGGER (una sola vez):
// 1. Activadores (ícono de reloj) → Agregar activador
// 2. Función: confirmarCompras
// 3. Evento: Por tiempo → Minutos → Cada 5 minutos
// 4. Guardar
function confirmarCompras() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  if (!ss) {
    Logger.log('No se encontró spreadsheet — verificá SPREADSHEET_ID');
    return;
  }
  var sheet = ss.getSheets()[0];
  var data = sheet.getDataRange().getValues();
  var header = data[0];

  // Encontrar columnas — usa includes() para matchear headers parciales
  // Headers reales: Marca temporal | Nombre | Email | Objetivo Principal | Suplemento 1 | Dosis 1 | ... | Tipo Compra | Precio Plan | Datos JSON
  var tipoCol = -1, emailCol = -1, nombreCol = -1, objetivoCol = -1;
  var sup1Col = -1, dosis1Col = -1, sup2Col = -1, dosis2Col = -1, sup3Col = -1, dosis3Col = -1;
  var precioCol = -1, jsonCol = -1;

  for (var c = 0; c < header.length; c++) {
    var h = String(header[c]).trim().toLowerCase();
    if (h === 'nombre') nombreCol = c;
    if (h === 'email') emailCol = c;
    if (h.includes('objetivo')) objetivoCol = c;
    if (h.includes('suplemento 1') || h === 'sup1') sup1Col = c;
    if (h.includes('dosis 1') || h === 'dosis1') dosis1Col = c;
    if (h.includes('suplemento 2') || h === 'sup2') sup2Col = c;
    if (h.includes('dosis 2') || h === 'dosis2') dosis2Col = c;
    if (h.includes('suplemento 3') || h === 'sup3') sup3Col = c;
    if (h.includes('dosis 3') || h === 'dosis3') dosis3Col = c;
    if (h.includes('tipo')) tipoCol = c;
    if (h.includes('precio')) precioCol = c;
    if (h.includes('json') || h.includes('datos')) jsonCol = c;
  }

  // Fallback por posición (col 0 = Marca temporal, datos empiezan en col 1)
  if (nombreCol === -1) nombreCol = 1;
  if (emailCol === -1) emailCol = 2;
  if (objetivoCol === -1) objetivoCol = 3;
  if (sup1Col === -1) sup1Col = 4;
  if (dosis1Col === -1) dosis1Col = 5;
  if (sup2Col === -1) sup2Col = 6;
  if (dosis2Col === -1) dosis2Col = 7;
  if (sup3Col === -1) sup3Col = 8;
  if (dosis3Col === -1) dosis3Col = 9;
  if (tipoCol === -1) tipoCol = 10;
  if (precioCol === -1) precioCol = 11;
  if (jsonCol === -1) jsonCol = 12;

  // Buscar columna de control "emails_enviados"
  var emailsEnviadosCol = -1;
  for (var c = 0; c < header.length; c++) {
    if (String(header[c]).trim().toLowerCase() === 'emails_enviados') {
      emailsEnviadosCol = c;
      break;
    }
  }
  // Si no existe la columna, crearla
  if (emailsEnviadosCol === -1) {
    emailsEnviadosCol = header.length;
    sheet.getRange(1, emailsEnviadosCol + 1).setValue('emails_enviados');
  }

  var confirmados = 0;

  for (var i = 1; i < data.length; i++) {
    var tipo = String(data[i][tipoCol] || '').trim().toLowerCase();
    var yaEnviado = String(data[i][emailsEnviadosCol] || '').trim().toLowerCase();

    if (tipo !== 'confirmado' || yaEnviado === 'si') continue;

    var nombre = String(data[i][nombreCol] || '').trim();
    var email = String(data[i][emailCol] || '').trim();
    var firstName = nombre.split(' ')[0] || 'Cliente';
    var objetivoKey = String(data[i][objetivoCol] || '');
    var objetivoNombre = objetivoNombreDisplay(objetivoKey);

    // Parsear JSON
    var jsonRaw = String(data[i][jsonCol] || '{}');
    var parsed = {};
    try { parsed = JSON.parse(jsonRaw); } catch(ex) {}

    var planNombre = parsed.plan || 'Plan Esencial';
    var frecuencia = parsed.frecuencia || 'mensual';
    var direccion = parsed.direccion || 'Pendiente de confirmar';

    // Precio
    var precio = '';
    if (parsed.precio && parsed.precio > 0) {
      precio = String(parsed.precio);
    } else {
      precio = String(data[i][precioCol] || '');
    }
    if (precio) {
      var num = parseInt(precio);
      if (!isNaN(num)) precio = num.toLocaleString('es-UY');
    }

    // Suplementos
    var sups = [];
    var s1 = String(data[i][sup1Col] || '');
    var d1 = String(data[i][dosis1Col] || '');
    var s2 = String(data[i][sup2Col] || '');
    var d2 = String(data[i][dosis2Col] || '');
    var s3 = String(data[i][sup3Col] || '');
    var d3 = String(data[i][dosis3Col] || '');
    if (s1) sups.push({ nombre: s1, dosis: d1 });
    if (s2) sups.push({ nombre: s2, dosis: d2 });
    if (s3) sups.push({ nombre: s3, dosis: d3 });

    // Extras del JSON
    if (parsed.suplementos && parsed.suplementos.length > 3) {
      for (var j = 3; j < parsed.suplementos.length; j++) {
        var match = parsed.suplementos[j].match(/^(.+?)\s*\((.+)\)$/);
        if (match) sups.push({ nombre: match[1].trim(), dosis: match[2].trim() });
        else sups.push({ nombre: parsed.suplementos[j].trim(), dosis: '' });
      }
    }

    if (!email || !email.includes('@')) continue;

    // Ignorar filas sin precio (leads puros que se marcaron confirmado por error)
    if (!precio || precio === '0') continue;

    // Enviar 1 email al cliente (confirmacion + plan) + 1 email interno
    var planHtml = '';
    try {
      planHtml = sendPlanEmail(email, firstName, objetivoNombre, sups, planNombre, frecuencia, precio, direccion) || '';
    } catch (planErr) {
      Logger.log('Error en sendPlanEmail (continua con notificacion): ' + planErr);
    }
    sendInternalNotification(nombre, email, planNombre, frecuencia, precio, sups, direccion, planHtml);

    // Marcar como enviado
    sheet.getRange(i + 1, emailsEnviadosCol + 1).setValue('si');
    confirmados++;

    Logger.log('Compra confirmada — emails enviados a: ' + email);
  }

  Logger.log('Total confirmados procesados: ' + confirmados);

  // También procesar envíos
  notificarEnvios();
}


// ===== EMAIL 3: Notificación de envío (automático cada 5 min) =====
// En el spreadsheet, cambiá "confirmado" por "enviado" cuando despachás el pedido.
// Esta función detecta filas con "enviado" y manda el email de "tu pedido está en camino".
function notificarEnvios() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  if (!ss) return;
  var sheet = ss.getSheets()[0];
  var data = sheet.getDataRange().getValues();
  var header = data[0];

  // Encontrar columnas
  var tipoCol = -1, emailCol = -1, nombreCol = -1;
  for (var c = 0; c < header.length; c++) {
    var h = String(header[c]).trim().toLowerCase();
    if (h === 'nombre') nombreCol = c;
    if (h === 'email') emailCol = c;
    if (h.includes('tipo')) tipoCol = c;
  }
  if (nombreCol === -1) nombreCol = 1;
  if (emailCol === -1) emailCol = 2;
  if (tipoCol === -1) tipoCol = 10;

  // Buscar columna "envio_notificado"
  var envioCol = -1;
  for (var c = 0; c < header.length; c++) {
    if (String(header[c]).trim().toLowerCase() === 'envio_notificado') {
      envioCol = c;
      break;
    }
  }
  if (envioCol === -1) {
    envioCol = header.length;
    sheet.getRange(1, envioCol + 1).setValue('envio_notificado');
  }

  for (var i = 1; i < data.length; i++) {
    var tipo = String(data[i][tipoCol] || '').trim().toLowerCase();
    var yaNotificado = String(data[i][envioCol] || '').trim().toLowerCase();

    if (tipo !== 'enviado' || yaNotificado === 'si') continue;

    var nombre = String(data[i][nombreCol] || '').trim();
    var email = String(data[i][emailCol] || '').trim();
    var firstName = nombre.split(' ')[0] || 'Cliente';

    if (!email || !email.includes('@')) continue;

    var html = '<div style="font-family:\'Helvetica Neue\',Arial,sans-serif; max-width:560px; margin:0 auto; background:#fff; border-radius:16px; overflow:hidden;">' +
      '<div style="padding:32px 40px 20px; text-align:center;">' +
      '<div style="font-size:17px; font-weight:600; color:#2E2E2E; letter-spacing:4px;">niva</div>' +
      '</div>' +
      '<div style="padding:0 40px 24px;">' +
      '<h1 style="font-size:26px; font-weight:300; color:#2E2E2E; margin:0 0 12px;">Tu pedido está en camino</h1>' +
      '<p style="font-size:15px; color:#4A4A4A; line-height:1.6;">' + firstName + ', ya enviamos tus suplementos. Deberían llegarte en las próximas 24-48 horas.</p>' +
      '</div>' +
      '<div style="padding:0 40px 24px;">' +
      '<div style="background:#F6F4F2; border-radius:12px; padding:20px;">' +
      '<div style="font-size:13px; font-weight:600; color:#6E766B; letter-spacing:1px; text-transform:uppercase; margin-bottom:10px;">Recordá</div>' +
      '<p style="font-size:14px; color:#4A4A4A; line-height:1.7; margin:0 0 6px;">· Revisá tu email para ver tu plan personalizado con dosis y horarios.</p>' +
      '<p style="font-size:14px; color:#4A4A4A; line-height:1.7; margin:0 0 6px;">· Empezá de a un suplemento por día para que tu cuerpo se adapte.</p>' +
      '<p style="font-size:14px; color:#4A4A4A; line-height:1.7; margin:0;">· En una semana te enviamos tu primer check-in.</p>' +
      '</div></div>' +
      '<div style="padding:0 40px 32px;">' +
      '<p style="font-size:14px; color:#4A4A4A; line-height:1.6;">¿Alguna duda? Escribinos a <a href="mailto:hola@niva.uy" style="color:#6E766B;">hola@niva.uy</a></p>' +
      '</div>' +
      '<div style="padding:24px 40px; background:#2E2E2E; text-align:center;">' +
      '<div style="font-size:14px; font-weight:600; color:#fff; letter-spacing:3px;">niva</div>' +
      '<div style="font-size:12px; color:rgba(255,255,255,0.6);">tu bienestar, más simple</div>' +
      '<div style="font-size:12px; color:rgba(255,255,255,0.4); margin-top:12px;">hola@niva.uy · niva.uy · <a href="https://instagram.com/nivauy" style="color:rgba(255,255,255,0.4);">@nivauy</a></div>' +
      '</div></div>';

    nivaSend(email, firstName + ', tu pedido niva está en camino', '', html);
    sheet.getRange(i + 1, envioCol + 1).setValue('si');
    Logger.log('Notificación de envío mandada a: ' + email);
  }
}


// ===== EMAIL 4: Check-in semanal =====
function checkinSemanal() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  if (!ss) return;
  const sheet = ss.getSheets()[0];
  const data = sheet.getDataRange().getValues();
  const header = data[0];

  const emailCol = header.indexOf('Email') !== -1 ? header.indexOf('Email') : 2;
  const nombreCol = header.indexOf('Nombre') !== -1 ? header.indexOf('Nombre') : 1;

  const thirtyDaysAgo = new Date();
  thirtyDaysAgo.setDate(thirtyDaysAgo.getDate() - 30);

  for (let i = 1; i < data.length; i++) {
    const timestamp = new Date(data[i][0]);
    if (timestamp > thirtyDaysAgo) {
      const email = data[i][emailCol];
      const nombre = String(data[i][nombreCol]).split(' ')[0];
      if (email && email.includes('@')) {
        enviarCheckin(email, nombre);
      }
    }
  }
}

function enviarCheckin(email, nombre) {
  const CHECKIN_FORM_URL = 'https://forms.gle/TU_FORM_CHECKIN'; // ← CAMBIAR

  const html = '<div style="font-family:\'Helvetica Neue\',Arial,sans-serif; max-width:560px; margin:0 auto; background:#fff; border-radius:16px; overflow:hidden;">' +
    '<div style="padding:32px 40px 20px; text-align:center;">' +
    '<div style="font-size:17px; font-weight:600; color:#2E2E2E; letter-spacing:4px;">niva</div>' +
    '</div>' +
    '<div style="padding:0 40px 24px;">' +
    '<h1 style="font-size:24px; font-weight:300; color:#2E2E2E; margin:0 0 12px;">¿Cómo te sentís esta semana?</h1>' +
    '<p style="font-size:15px; color:#4A4A4A; line-height:1.6;">' + nombre + ', son solo 5 preguntas rápidas (1 minuto). Tu feedback nos ayuda a ajustar tu plan.</p>' +
    '</div>' +
    '<div style="padding:0 40px 32px; text-align:center;">' +
    '<a href="' + CHECKIN_FORM_URL + '" style="display:inline-block; padding:14px 40px; background:#2E2E2E; color:#fff; border-radius:24px; text-decoration:none; font-size:15px; font-weight:500;">Completar check-in</a>' +
    '</div>' +
    '<div style="padding:24px 40px; background:#2E2E2E; text-align:center;">' +
    '<div style="font-size:14px; font-weight:600; color:#fff; letter-spacing:3px;">niva</div>' +
    '<div style="font-size:12px; color:rgba(255,255,255,0.6);">tu bienestar, más simple</div>' +
    '</div></div>';

  nivaSend(email, nombre + ', ¿cómo vas con tu plan?', '', html);
}


// ===== EMAIL 5: Follow-up para leads que no compraron =====
// Configurar: Activadores → Nuevo → followUpLeads → Por tiempo → Diario (cada 24hs)
function followUpLeads() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  if (!ss) return;
  var sheet = ss.getSheets()[0];
  var data = sheet.getDataRange().getValues();
  var header = data[0];

  // Encontrar columnas
  var tipoCol = -1;
  var emailCol = -1;
  var nombreCol = -1;
  var objetivoCol = -1;
  var jsonCol = -1;
  var sup1Col = -1;
  var sup2Col = -1;
  var sup3Col = -1;

  for (var c = 0; c < header.length; c++) {
    var h = String(header[c]).trim().toLowerCase();
    if (h === 'nombre') nombreCol = c;
    if (h === 'email') emailCol = c;
    if (h.includes('objetivo')) objetivoCol = c;
    if (h.includes('tipo')) tipoCol = c;
    if (h.includes('json') || h.includes('datos')) jsonCol = c;
    if (h.includes('suplemento 1') || h === 'sup1') sup1Col = c;
    if (h.includes('suplemento 2') || h === 'sup2') sup2Col = c;
    if (h.includes('suplemento 3') || h === 'sup3') sup3Col = c;
  }

  // Fallback por posición (col 0 = Marca temporal)
  if (nombreCol === -1) nombreCol = 1;
  if (emailCol === -1) emailCol = 2;
  if (objetivoCol === -1) objetivoCol = 3;
  if (sup1Col === -1) sup1Col = 4;
  if (sup2Col === -1) sup2Col = 6;
  if (sup3Col === -1) sup3Col = 8;
  if (tipoCol === -1) tipoCol = 10;
  if (jsonCol === -1) jsonCol = 12;

  // Check leads from 2-48 hours ago that haven't purchased
  var now = new Date();
  var twoHoursAgo = new Date(now.getTime() - 2 * 60 * 60 * 1000);
  var fortyEightHoursAgo = new Date(now.getTime() - 48 * 60 * 60 * 1000);

  // Build set of emails that DID purchase (to exclude from follow-up)
  var purchasedEmails = {};
  for (var i = 1; i < data.length; i++) {
    var tipo = String(data[i][tipoCol] || '').trim();
    if (tipo !== 'pendiente' && tipo !== '' && tipo !== '0') {
      purchasedEmails[String(data[i][emailCol]).trim().toLowerCase()] = true;
    }
  }

  // Build set of emails already followed up (check if 'seguimiento_enviado' column exists)
  var followUpCol = header.indexOf('seguimiento_enviado');

  for (var i = 1; i < data.length; i++) {
    var timestamp = new Date(data[i][0]);
    var tipo = String(data[i][tipoCol] || '').trim();
    var email = String(data[i][emailCol] || '').trim();
    var nombre = String(data[i][nombreCol] || '').trim().split(' ')[0];

    // Only leads (pendiente) from 2-48 hours ago, not yet purchased, not yet followed up
    if (tipo !== 'pendiente') continue;
    if (timestamp > twoHoursAgo || timestamp < fortyEightHoursAgo) continue;
    if (!email || !email.includes('@')) continue;
    if (purchasedEmails[email.toLowerCase()]) continue;
    if (followUpCol !== -1 && data[i][followUpCol] === 'si') continue;

    // Parse their supplements
    var sups = [];
    var s1 = String(data[i][sup1Col] || '');
    var s2 = String(data[i][sup2Col] || '');
    var s3 = String(data[i][sup3Col] || '');
    if (s1) sups.push(s1);
    if (s2) sups.push(s2);
    if (s3) sups.push(s3);

    var objetivoKey = String(data[i][objetivoCol] || '');
    var objetivoNombre = objetivoNombreDisplay(objetivoKey);

    // Send follow-up
    sendFollowUpEmail(email, nombre, objetivoNombre, sups);

    // Mark as followed up
    if (followUpCol !== -1) {
      sheet.getRange(i + 1, followUpCol + 1).setValue('si');
    }

    Logger.log('Follow-up enviado a: ' + email);
  }
}


function sendFollowUpEmail(email, nombre, objetivo, supNames) {
  // Build supplement preview
  var supsPreview = '';
  supNames.forEach(function(supName) {
    var color = buscarEnMapeo(DISC_COLORS, supName) || '#B8C1B6';
    var label = buscarEnMapeo(DISC_LABELS, supName) || '•';
    var ciencia = buscarEnMapeo(CIENCIA, supName) || {};
    supsPreview += '<div style="margin-bottom:8px;">' +
      '<table cellpadding="0" cellspacing="0"><tr>' +
      '<td style="width:36px; height:36px; border-radius:50%; background:' + color + '; text-align:center; vertical-align:middle; font-weight:600; font-size:13px; color:rgba(46,46,46,0.55); line-height:36px;">' + label + '</td>' +
      '<td style="padding-left:10px; vertical-align:middle;">' +
      '<div style="font-size:14px; font-weight:500; color:#2E2E2E;">' + supName + '</div>' +
      '<div style="font-size:12px; color:#7A7972;">' + (ciencia.tiempo || '') + '</div>' +
      '</td></tr></table></div>';
  });

  var html = '<div style="font-family:\'Helvetica Neue\',Arial,sans-serif; max-width:560px; margin:0 auto; background:#fff; border-radius:16px; overflow:hidden;">' +
    '<div style="padding:32px 40px 20px; text-align:center;">' +
    '<div style="font-size:17px; font-weight:600; color:#2E2E2E; letter-spacing:4px;">niva</div>' +
    '</div>' +
    '<div style="padding:0 40px 24px;">' +
    '<h1 style="font-size:24px; font-weight:300; color:#2E2E2E; margin:0 0 8px;">' + nombre + ', tu plan te está esperando</h1>' +
    '<p style="font-size:15px; color:#4A4A4A; line-height:1.6;">Armamos una selección personalizada para ' + (objetivo.includes(' y ') ? 'tus objetivos de' : 'tu objetivo de') + ' <strong>' + objetivo.toLowerCase() + '</strong>. Estos suplementos fueron elegidos con respaldo científico para vos:</p>' +
    '</div>' +
    '<div style="padding:0 40px 24px;">' +
    '<div style="background:#F6F4F2; border-radius:12px; padding:20px;">' +
    '<div style="font-size:13px; font-weight:600; color:#6E766B; letter-spacing:1px; text-transform:uppercase; margin-bottom:12px;">Tu selección</div>' +
    supsPreview +
    '</div></div>' +
    '<div style="padding:0 40px 24px;">' +
    '<p style="font-size:14px; color:#4A4A4A; line-height:1.6;">Los efectos empiezan a notarse entre 2 y 6 semanas. Cuanto antes empieces, antes sentís la diferencia.</p>' +
    '</div>' +
    '<div style="padding:0 40px 32px; text-align:center;">' +
    '<a href="https://niva.uy" style="display:inline-block; padding:14px 40px; background:#2E2E2E; color:#fff; border-radius:24px; text-decoration:none; font-size:15px; font-weight:500;">Empezar mi plan</a>' +
    '</div>' +
    '<div style="padding:24px 40px; background:#2E2E2E; text-align:center;">' +
    '<div style="font-size:14px; font-weight:600; color:#fff; letter-spacing:3px;">niva</div>' +
    '<div style="font-size:12px; color:rgba(255,255,255,0.6);">tu bienestar, más simple</div>' +
    '<div style="font-size:12px; color:rgba(255,255,255,0.4); margin-top:12px;">hola@niva.uy · niva.uy · <a href="https://instagram.com/nivauy" style="color:rgba(255,255,255,0.4);">@nivauy</a></div>' +
    '</div></div>';

  nivaSend(email, nombre + ', tu plan personalizado te espera', '', html);
}


// ===== REENVIAR EMAILS FALLIDOS =====
// Uso: corregir el email en el spreadsheet, poner "reenviar" en la columna emails_enviados, y ejecutar esta funcion.
// Busca automaticamente todas las filas marcadas con "reenviar" y les manda el email que corresponda.
function reenviarFallidos() {
  var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Respuestas') ||
    SpreadsheetApp.openById(SPREADSHEET_ID).getSheets()[0];
  var data = sheet.getDataRange().getValues();
  var header = data[0];

  // Buscar columnas (misma logica que confirmarCompras)
  var nombreCol = -1, emailCol = -1, objetivoCol = -1, jsonCol = -1, tipoCol = -1, precioCol = -1, emailsEnviadosCol = -1;
  var sup1Col = -1, dosis1Col = -1, sup2Col = -1, dosis2Col = -1, sup3Col = -1, dosis3Col = -1;
  for (var c = 0; c < header.length; c++) {
    var h = String(header[c]).trim().toLowerCase();
    if (h === 'nombre') nombreCol = c;
    if (h === 'email') emailCol = c;
    if (h.includes('objetivo')) objetivoCol = c;
    if (h.includes('suplemento 1') || h === 'sup1') sup1Col = c;
    if (h.includes('dosis 1') || h === 'dosis1') dosis1Col = c;
    if (h.includes('suplemento 2') || h === 'sup2') sup2Col = c;
    if (h.includes('dosis 2') || h === 'dosis2') dosis2Col = c;
    if (h.includes('suplemento 3') || h === 'sup3') sup3Col = c;
    if (h.includes('dosis 3') || h === 'dosis3') dosis3Col = c;
    if (h.includes('tipo')) tipoCol = c;
    if (h.includes('precio')) precioCol = c;
    if (h.includes('json') || h.includes('datos')) jsonCol = c;
    if (h === 'emails_enviados') emailsEnviadosCol = c;
  }

  // Fallback por posicion (igual que confirmarCompras)
  if (nombreCol === -1) nombreCol = 1;
  if (emailCol === -1) emailCol = 2;
  if (objetivoCol === -1) objetivoCol = 3;
  if (sup1Col === -1) sup1Col = 4;
  if (dosis1Col === -1) dosis1Col = 5;
  if (sup2Col === -1) sup2Col = 6;
  if (dosis2Col === -1) dosis2Col = 7;
  if (sup3Col === -1) sup3Col = 8;
  if (dosis3Col === -1) dosis3Col = 9;
  if (tipoCol === -1) tipoCol = 10;

  if (emailsEnviadosCol === -1) {
    Logger.log('No se encontro la columna emails_enviados. Creala primero.');
    return;
  }

  var reenviados = 0;

  for (var i = 1; i < data.length; i++) {
    var marcador = String(data[i][emailsEnviadosCol] || '').trim().toLowerCase();
    if (marcador !== 'reenviar') continue;

    var fila = i + 1; // numero de fila en Sheets (1-indexed + header)
    var nombre = String(data[i][nombreCol] || '').trim();
    var email = String(data[i][emailCol] || '').trim().replace(/\s+/g, '').toLowerCase();
    var firstName = nombre.split(' ')[0] || 'Cliente';
    var objetivoKey = String(data[i][objetivoCol] || '');
    var objetivoNombre = objetivoNombreDisplay(objetivoKey);

    if (!email || !email.includes('@')) {
      Logger.log('Fila ' + fila + ': email invalido "' + email + '". Corregi el email primero.');
      continue;
    }

    // Parsear JSON
    var jsonRaw = jsonCol >= 0 ? String(data[i][jsonCol] || '{}') : '{}';
    var parsed = {};
    try { parsed = JSON.parse(jsonRaw); } catch(ex) {}

    // Armar sups: primero de columnas individuales, luego del JSON como fallback
    var sups = [];
    var s1 = sup1Col >= 0 ? String(data[i][sup1Col] || '').trim() : '';
    var d1 = dosis1Col >= 0 ? String(data[i][dosis1Col] || '').trim() : '';
    var s2 = sup2Col >= 0 ? String(data[i][sup2Col] || '').trim() : '';
    var d2 = dosis2Col >= 0 ? String(data[i][dosis2Col] || '').trim() : '';
    var s3 = sup3Col >= 0 ? String(data[i][sup3Col] || '').trim() : '';
    var d3 = dosis3Col >= 0 ? String(data[i][dosis3Col] || '').trim() : '';
    if (s1) sups.push({ nombre: s1, dosis: d1 });
    if (s2) sups.push({ nombre: s2, dosis: d2 });
    if (s3) sups.push({ nombre: s3, dosis: d3 });

    // Si no habia en columnas individuales, intentar del JSON
    if (sups.length === 0 && parsed.suplementos && parsed.suplementos.length > 0) {
      for (var j = 0; j < parsed.suplementos.length; j++) {
        var match = parsed.suplementos[j].match(/^(.+?)\s*\((.+?)\)$/);
        if (match) sups.push({ nombre: match[1].trim(), dosis: match[2].trim() });
        else sups.push({ nombre: parsed.suplementos[j].trim(), dosis: '' });
      }
    }
    // Extras del JSON (sups 4+)
    if (parsed.suplementos && parsed.suplementos.length > 3) {
      for (var j = 3; j < parsed.suplementos.length; j++) {
        var match = parsed.suplementos[j].match(/^(.+?)\s*\((.+?)\)$/);
        if (match) sups.push({ nombre: match[1].trim(), dosis: match[2].trim() });
        else sups.push({ nombre: parsed.suplementos[j].trim(), dosis: '' });
      }
    }

    if (sups.length === 0) {
      Logger.log('Fila ' + fila + ': no se encontraron suplementos en columnas ni JSON. Verifica los datos.');
      continue;
    }

    // Determinar que tipo de email reenviar
    var tipo = tipoCol >= 0 ? String(data[i][tipoCol] || '').trim().toLowerCase() : '';
    var precio = precioCol >= 0 ? String(data[i][precioCol] || '') : '';
    var hasPrecio = precio && precio !== '0' && precio !== '';

    try {
      if (tipo === 'confirmado') {
        // Compra confirmada — reenviar plan + confirmacion
        var planNombre = parsed.plan || 'Plan Esencial';
        var frecuencia = parsed.frecuencia || 'mensual';
        var direccion = parsed.direccion || 'Pendiente de confirmar';
        Logger.log('Fila ' + fila + ': reenviando email de COMPRA CONFIRMADA a ' + email);
        sendPlanEmail(email, firstName, objetivoNombre, sups, planNombre, frecuencia, precio, direccion);
      } else if (tipo === 'pendiente' && hasPrecio) {
        // Pendiente de pago — reenviar email de pendiente con deep link
        var planNombre = parsed.plan || 'Plan Esencial';
        var frecuencia = parsed.frecuencia || 'mensual';
        Logger.log('Fila ' + fila + ': reenviando email de PENDIENTE a ' + email);
        sendPendingEmail(email, firstName, planNombre, frecuencia, precio, sups, objetivoNombre);
      } else {
        // Lead — reenviar recomendacion
        Logger.log('Fila ' + fila + ': reenviando email de LEAD a ' + email);
        sendLeadEmail(email, firstName, objetivoNombre, sups);
      }

      // Marcar como enviado
      sheet.getRange(fila, emailsEnviadosCol + 1).setValue('si');
      reenviados++;
      Logger.log('Fila ' + fila + ': reenviado OK a ' + email);
    } catch (err) {
      Logger.log('Fila ' + fila + ': ERROR — ' + err.toString());
    }
  }

  Logger.log('Reenvio completado. Total reenviados: ' + reenviados);
}


// ===== TEST: para verificar que los emails salen desde hola@niva.uy =====
// Ejecutar manualmente para forzar la autorización de GmailApp
function testEnvio() {
  nivaSend('alejandrabotta@gmail.com', 'test niva', '', '<p>test desde hola@niva.uy</p>');
}

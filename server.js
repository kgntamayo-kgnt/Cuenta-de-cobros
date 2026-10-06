import express from 'express';
import path from 'path';
import { fileURLToPath } from 'url';
import { GoogleGenAI, Type } from '@google/genai';
import mammoth from 'mammoth';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const app = express();
const PORT = process.env.PORT || 3000;

app.use(express.json({ limit: '50mb' }));
app.use(express.static(__dirname));

// Serve client-side mammoth library
app.get('/mammoth.browser.min.js', (req, res) => {
  res.sendFile(path.join(__dirname, 'node_modules/mammoth/mammoth.browser.min.js'));
});

function cleanWordText(raw) {
  if (!raw) return '';
  const rawLines = raw.split(/\r?\n/);
  const lines = [];
  for (const rl of rawLines) {
    if (rl.includes('\t')) {
      rl.split('\t').forEach(sub => lines.push(sub));
    } else {
      lines.push(rl);
    }
  }

  const cleanList = [];
  const seen = new Set();
  const ignorePatterns = [
    /^\s*informe\s+mensual/i,
    /^\s*formato\s+informe/i,
    /^\s*ministerio\s+de\s+minas/i,
    /^\s*contrato\s+ggc/i,
    /^\s*contratista:/i,
    /^\s*supervisor:/i,
    /^\s*objeto\s+del\s+contrato/i,
    /^\s*obligaciones\s+y\/o\s+actividades/i,
    /^\s*avances\s+y\s+logros/i,
    /^\s*documentos\s+empleados\s+o\s+desarrollados/i,
    /^\s*no\.\s+obligaciones/i,
    /^\s*\(no\s+se\s+registraron\s+evidencias/i,
    /^\s*\(sin\s+evidencias/i,
    /^\s*reporte\s+de\s+producciones/i,
    /^\s*desglose/i,
    /^\s*reuniones\s+y\s+eventos/i,
    /^\s*no\s+se\s+encontró\s+archivo/i
  ];

  for (let line of lines) {
    line = line.trim();
    if (!line) continue;
    line = line.replace(/^[\s•\-\*\u2022\u25E6\u25AA\u25AB\u2013\u2014]+/, '').trim();
    if (/^\d+[\.\)]\s+/.test(line) && !/^\d-202/i.test(line)) {
      line = line.replace(/^\d+[\.\)]\s+/, '').trim();
    }
    if (!line) continue;

    // Has a file extension (like .pdf, .docx, .zip) or is a radicado (1-202..., 2-202..., 3-202...) or is a folder
    const hasFileExt = /\.[a-z0-9]{2,5}$/i.test(line);
    const isRadicado = /^(1|2|3)-202/i.test(line) || /^\d{14,18}/.test(line);
    const isFolder = /^carpeta:\s*/i.test(line);

    // If it's a file, radicado or folder, keep it unconditionally!
    if (hasFileExt || isRadicado || isFolder) {
      const key = line.toLowerCase();
      if (!seen.has(key)) {
        seen.add(key);
        cleanList.push(line);
      }
      continue;
    }

    if (ignorePatterns.some(p => p.test(line))) continue;
    if (/^obligaci[oó]n\s+\d/i.test(line)) continue;

    const key = line.toLowerCase();
    if (!seen.has(key)) {
      seen.add(key);
      cleanList.push(line);
    }
  }

  // If cleanList somehow ended up empty, fallback to raw so nothing is lost
  if (cleanList.length === 0) {
    return raw.trim();
  }

  return cleanList.join('\n');
}

// Endpoint to extract text from Word documents (.docx)
app.post('/api/parse-docx', async (req, res) => {
  try {
    const { base64 } = req.body;
    if (!base64) {
      return res.status(400).json({ error: 'No se envió contenido de archivo en base64.' });
    }
    const buffer = Buffer.from(base64, 'base64');
    const result = await mammoth.extractRawText({ buffer });
    const rawVal = result.value || '';
    const cleaned = cleanWordText(rawVal);
    return res.json({ 
      text: cleaned || rawVal,
      rawText: rawVal 
    });
  } catch (err) {
    console.error('Error procesando archivo DOCX:', err);
    return res.status(500).json({ error: 'Error al extraer texto del documento Word: ' + err.message });
  }
});

// Initialize Gemini SDK server-side
let ai = null;
if (process.env.GEMINI_API_KEY) {
  try {
    ai = new GoogleGenAI({
      apiKey: process.env.GEMINI_API_KEY,
      httpOptions: {
        headers: {
          'User-Agent': 'aistudio-build',
        },
      },
    });
  } catch (err) {
    console.warn('[Gemini Init Warning]', err.message);
  }
}

// Fallback deterministic rule classifier with custom rules support
function classifyByRules(name, folderContext = '', customRules = []) {
  const lower = (name + ' ' + folderContext).toLowerCase();

  // Evaluamos primero las reglas personalizadas del usuario si existen
  if (Array.isArray(customRules)) {
    for (const rule of customRules) {
      if (rule.enabled === false) continue;
      const pat = (rule.pattern || '').toLowerCase().trim();
      if (!pat) continue;
      const obId = parseInt(rule.obligationId, 10);
      if (isNaN(obId) || obId < 1 || obId > 7) continue;

      if (rule.matchType === 'startsWith' && lower.startsWith(pat)) {
        return { obligationId: obId, reason: rule.description || `Regla personalizada: inicia con "${rule.pattern}"` };
      }
      if (rule.matchType === 'contains' && lower.includes(pat)) {
        return { obligationId: obId, reason: rule.description || `Regla personalizada: contiene "${rule.pattern}"` };
      }
      if (rule.matchType === 'endsWith' && lower.endsWith(pat)) {
        return { obligationId: obId, reason: rule.description || `Regla personalizada: termina en "${rule.pattern}"` };
      }
    }
  }

  // Regla 6: Actas y respuestas legislativas / Congreso / Derechos de Petición
  if (
    lower.includes('rta dp') ||
    lower.includes('gobernanza resiliente') ||
    lower.includes('respuesta hr') ||
    lower.includes('legislativo') ||
    lower.includes('ana_erazo') ||
    lower.includes('acta posesión') ||
    lower.includes('posesion') ||
    lower.includes('ley_2603') ||
    lower.includes('circular_din') ||
    lower.includes('pronter') ||
    lower.includes('derecho de petición') ||
    lower.includes('senado') ||
    lower.includes('congreso')
  ) {
    return { obligationId: 6, reason: 'Actas de reunión y respuestas al poder legislativo o peticiones ciudadanas' };
  }

  // Regla 2: Instalaciones, licencias, inspecciones técnicas y radicados 1-202..., 2-202..., 3-202...
  if (
    lower.startsWith('1-202') ||
    lower.startsWith('2-202') ||
    lower.startsWith('3-202') ||
    lower.startsWith('120264') ||
    lower.startsWith('20264005') ||
    lower.includes('1-2026-') ||
    lower.includes('2-2026-') ||
    lower.includes('3-2026-') ||
    lower.includes('fuente gdr') ||
    lower.includes('fuente sr') ||
    lower.includes('fuente sr--am') ||
    lower.includes('instructivo de acceso') ||
    lower.includes('instructivo acceso') ||
    lower.includes('intructivo acceso') ||
    lower.includes('f-tnu-mrn') ||
    lower.includes('in-tnu-mrn') ||
    lower.includes('calderon') ||
    lower.includes('calificaciones reactor') ||
    lower.includes('examen responsable de mantenimiento') ||
    lower.includes('disponibilida personal') ||
    lower.includes('examen camilo') ||
    lower.includes('licencia') ||
    lower.includes('radicado')
  ) {
    return { obligationId: 2, reason: 'Radicado oficial o trámite de instalaciones radiactivas/nucleares (1-202... / 2-202...)' };
  }

  // Regla 3: TdRs, Contratos, Convenios, Propuestas BID, ANSN
  if (
    lower.includes('tdr') ||
    lower.includes('propfin') ||
    lower.includes('propuestafinaciación') ||
    lower.includes('propuestafinanciacion') ||
    lower.includes('consultor individual') ||
    lower.includes('prioridades macro bid') ||
    lower.includes('propuesta  discusion bid') ||
    lower.includes('propuesta discusion bid') ||
    lower.includes('carpeta: tdr-bid') ||
    lower.includes('carpeta: 01-fortalecimiento ansn') ||
    lower.includes('carpeta: 02-nucleoelectricidad') ||
    lower.includes('presentacion bid') ||
    lower.includes('convenio') ||
    lower.includes('contratacion')
  ) {
    return { obligationId: 3, reason: 'Seguimiento a contratos, convenios y TdRs (BID / ANSN / Consultores)' };
  }

  // Regla 5: NORM, mesas técnicas, cooperación internacional, presentaciones
  if (
    lower.includes('norm') ||
    lower.includes('argentina') ||
    lower.includes('brasil') ||
    lower.includes('españa') ||
    lower.includes('espana') ||
    lower.includes('cnen') ||
    lower.includes('ansn') ||
    lower.includes('csn') ||
    lower.includes('fosfoyeso') ||
    lower.includes('mise_curso') ||
    lower.includes('mise_ejemplo') ||
    lower.includes('presentacion sgc') ||
    lower.includes('presentacion norm') ||
    lower.includes('presentación institucional norm') ||
    lower.includes('cooperacion-mme-bid-oiea') ||
    lower.includes('mesa técnica') ||
    lower.includes('mesa tecnica')
  ) {
    return { obligationId: 5, reason: 'Mesas técnicas de trabajo, NORM/TENORM y cooperación técnica internacional' };
  }

  // Regla 1: Hitos OIEA, NEPIO, Política Pública, Hojas de Ruta
  if (
    lower.includes('nepio') ||
    lower.includes('nuclear power programme') ||
    lower.includes('enfoque_de_hitos') ||
    lower.includes('enfoquede hitos') ||
    lower.includes('fase1_enfoque') ||
    lower.includes('fase1_limites') ||
    lower.includes('iaea_milestone') ||
    lower.includes('memorando estratégico') ||
    lower.includes('memorando estrategico') ||
    lower.includes('doe') ||
    lower.includes('taxonomía verde') ||
    lower.includes('nucleoelectricidad') ||
    lower.includes('nuclear_financial_strategy') ||
    lower.includes('nuclear_energy_evaluation') ||
    lower.includes('olacde') ||
    lower.includes('politica publica') ||
    lower.includes('política pública')
  ) {
    return { obligationId: 1, reason: 'Política pública de nucleoelectricidad y metodología de Hitos OIEA' };
  }

  // Regla 4: Mitigación de riesgos, hallazgos, inspección y evidencias fotográficas
  if (
    lower.includes('anexos') ||
    lower.includes('mitigacion') ||
    lower.includes('mitigación') ||
    lower.includes('riesgos') ||
    lower.includes('inspección') ||
    lower.includes('inspeccion') ||
    lower.includes('8833896c') ||
    lower.includes('c916eaf1') ||
    lower.endsWith('.jpg') ||
    lower.endsWith('.jpeg') ||
    lower.endsWith('.png')
  ) {
    return { obligationId: 4, reason: 'Informes técnicos de mitigación de riesgos y evidencias gráficas' };
  }

  // Regla 7: Soporte normativo general, PND, VPN/FortiClient, cuentas de cobro, formatos MME
  if (
    lower.includes('borrador r626') ||
    lower.includes('borrador r763') ||
    lower.includes('nuevo pnd') ||
    lower.includes('aportes gan al pnd') ||
    lower.includes('palabras apertura') ||
    lower.includes('firma-minenergia') ||
    lower.includes('gc-m2-f24') ||
    lower.includes('presentación gan') ||
    lower.includes('plantiila') ||
    lower.includes('plantilla') ||
    lower.includes('plant1') ||
    lower.includes('septiembre.zip') ||
    lower.includes('forticlient') ||
    lower.includes('vpn') ||
    lower.includes('index.html') ||
    lower.includes('cuenta de cobro') ||
    lower.includes('cuenta cobro')
  ) {
    return { obligationId: 7, reason: 'Actividades inherentes, administrativas, normativas generales y formatos MME' };
  }

  return { obligationId: 2, reason: 'Asignado a gestión técnica de trámites/instalaciones' };
}

// Balanceador para garantizar que NINGUNA de las 7 obligaciones quede vacía (Cobertura Total)
function balanceObligationsCoverage(items, classifications) {
  if (!Array.isArray(items) || items.length < 7) {
    return classifications;
  }

  const counts = { 1: 0, 2: 0, 3: 0, 4: 0, 5: 0, 6: 0, 7: 0 };
  classifications.forEach(c => {
    if (counts[c.obligationId] !== undefined) {
      counts[c.obligationId]++;
    }
  });

  const emptyObligations = [1, 2, 3, 4, 5, 6, 7].filter(id => counts[id] === 0);
  if (emptyObligations.length === 0) {
    return classifications; // Cobertura 7 de 7 completa
  }

  const keywordsByOb = {
    1: ['hitos', 'estrateg', 'politica', 'marco', 'plan', 'doe', 'nepio', 'evaluacion', 'nuclear', 'energia', 'roadmap', 'taxonom'],
    2: ['instalac', 'tramite', 'soporte', 'licencia', 'inspecc', 'calific', 'tecnico', 'expediente', 'acceso', 'reactor', 'radicado', 'fuente'],
    3: ['propuesta', 'seguimiento', 'convenio', 'consultor', 'contrat', 'ansn', 'bid', 'financ', 'prioridad', 'tdr'],
    4: ['riesgo', 'mitigac', 'inspecc', 'hallazgo', 'anexo', 'foto', 'evidencia', 'tecnico', 'fuente', 'fotogr'],
    5: ['cooperac', 'mesa', 'norm', 'internacional', 'presentacion', 'reunion', 'foro', 'taller', 'comite', 'bilateral'],
    6: ['acta', 'compromiso', 'respuesta', 'legislativ', 'circular', 'comunicacion', 'resumen', 'pronter', 'senado', 'congreso', 'peticion', 'ley'],
    7: ['general', 'administrativ', 'reunion', 'soporte', 'induccion', 'formato', 'actividad', 'cuenta', 'plantilla', 'vpn', 'forticlient', 'pnd']
  };

  emptyObligations.forEach(targetOb => {
    let bestIndex = -1;
    let bestScore = -1;

    for (let i = 0; i < items.length; i++) {
      const currentOb = classifications[i].obligationId;
      // Solo tomamos de obligaciones que tengan más de 1 documento
      if (counts[currentOb] > 1) {
        const lowerName = (items[i].name + ' ' + (items[i].folderContext || '')).toLowerCase();
        const kwList = keywordsByOb[targetOb] || [];
        let score = 0;
        kwList.forEach(k => {
          if (lowerName.includes(k)) score += 3;
        });

        // Dar prioridad a tomar de obligaciones con exceso de documentos
        if (counts[currentOb] > 5) score += 2;
        else if (counts[currentOb] > 2) score += 1;

        // Si el destino es la 7, solo transferir si tiene afinidad administrativa/general
        if (targetOb === 7 && (lowerName.includes('administrativ') || lowerName.includes('reunion') || lowerName.includes('cuenta') || lowerName.includes('soporte') || lowerName.includes('formato') || lowerName.includes('pnd') || lowerName.includes('vpn'))) {
          score += 4;
        }

        if (score > bestScore) {
          bestScore = score;
          bestIndex = i;
        }
      }
    }

    if (bestIndex !== -1) {
      const oldOb = classifications[bestIndex].obligationId;
      counts[oldOb]--;
      counts[targetOb]++;
      classifications[bestIndex] = {
        obligationId: targetOb,
        reason: `Asignado para garantizar cobertura total de la Obligación ${targetOb} (${classifications[bestIndex].reason || 'documento afín'})`
      };
    }
  });

  return classifications;
}

// Server API to classify with Gemini or Fallback
app.post('/api/classify', async (req, res) => {
  const { items, customRules = [] } = req.body;

  if (!items || !Array.isArray(items) || items.length === 0) {
    return res.status(400).json({ error: 'No se suministró una lista válida de archivos o carpetas.' });
  }

  // Build custom rules prompt section if user provided custom rules
  let customRulesPrompt = '';
  if (Array.isArray(customRules) && customRules.length > 0) {
    const activeRules = customRules.filter(r => r.enabled !== false);
    if (activeRules.length > 0) {
      customRulesPrompt = `\nREGLAS ESPECÍFICAS ADICIONALES CONFIGURADAS POR EL USUARIO (MÁXIMA PRIORIDAD):\n` +
        activeRules.map((r, i) => `${i + 1}. Si el nombre del documento ${r.matchType === 'startsWith' ? 'comienza con' : r.matchType === 'endsWith' ? 'termina con' : 'contiene'} "${r.pattern}" -> Asignar estrictamente a la Obligación ${r.obligationId} (Motivo: ${r.description || 'Regla personalizada'}).`).join('\n') + '\n';
    }
  }

  // Try Gemini if available
  if (ai) {
    try {
      const itemsPayload = items.map((it, idx) => ({
        index: idx,
        name: it.name,
        isFolder: !!it.isFolder,
        folderContext: it.folderContext || '',
      }));

      const prompt = `Eres un experto asistente técnico contractual para el Grupo de Asuntos Nucleares (GAN) del Ministerio de Minas y Energía de Colombia (MME).
Debes clasificar cada uno de los archivos y carpetas del listado mensual en EXACTAMENTE UNA de las 7 obligaciones contractuales estándar, SIN REPETICIÓN (cada documento aparece solo una vez).

REGLAS DE ORO OBLIGATORIAS:
1. COBERTURA TOTAL DEL INFORME: NINGUNA de las 7 obligaciones puede quedar con 0 documentos. El informe de supervisión exige que TODAS y cada una de las 7 obligaciones cuente con documentos de soporte. Si una categoría tiene pocas evidencias directas, asigna documentos o producciones afines para que las 7 queden cubiertas.
2. PRIORIZACIÓN TÉCNICA (OBLIGACIONES 1 A 6 vs OBLIGACIÓN 7):
   - La Obligación 7 es de carácter residual y estrictamente administrativo o no técnico (reuniones generales administrativas, inducción VPN/FortiClient, formatos del ministerio, cuentas de cobro).
   - DEBES PRIORIZAR ENÉRGICAMENTE el contenido en las Obligaciones 1 a 6 para todo documento técnico, regulatorio, de política, de supervisión contractual o legislativo.
   - NO envíes a la Obligación 7 ningún archivo que tenga contenido técnico correspondiente a las Obligaciones 1, 2, 3, 4, 5 o 6.

LAS 7 OBLIGACIONES CONTRACTUALES ESTÁNDAR:
1: Obligación 1 - Acuerdos e Hitos Nucleares / Política Pública (Hitos OIEA Milestone Approach, NEPIO, Hojas de Ruta, nucleoelectricidad, reactores, convenios de política energética con DOE u organismos).
2: Obligación 2 - Apoyo Técnico en Instalaciones, Licencias e Inspecciones (REGLA FUNDAMENTAL: cualquier archivo o carpeta que empiece por 1-202..., 2-202..., 3-202..., 120264..., 20264... o que haga referencia a fuentes GDR, Sr, Am, instructivos de acceso, exámenes de operadores de reactor, radicados y licencias de instalaciones nucleares o radiactivas).
3: Obligación 3 - Seguimiento a Contratos, Convenios y TdRs (Términos de Referencia TdR, contratación y propuestas con BID, fortalecimiento ANSN, consultor individual, seguimiento técnico-administrativo de convenios).
4: Obligación 4 - Informes Técnicos de Mitigación de Riesgos y Fuentes (Informes técnicos de mitigación, hallazgos de inspección, evaluación de fuentes huérfanas/en desuso, registros fotográficos o anexos de soporte de inspección).
5: Obligación 5 - Articulación de Mesas Técnicas, Cooperación y NORM (Material radiactivo natural NORM/TENORM, fosfoyeso, minería, O&G, mesas de trabajo técnicas y presentaciones de cooperación internacional con Argentina, Brasil, España, BID, OIEA, OLACDE, SGC).
6: Obligación 6 - Elaboración y Revisión de Actas y Respuestas Legislativas (Actas de reuniones de seguimiento o posesión, respuestas a Derechos de Petición DP, respuestas a congresistas/senadores/Cámara de Representantes, proyectos de ley).
7: Obligación 7 - Demás Actividades Inherentes / Normativa / Administrativas (Actividades de soporte general administrativo no técnico, reuniones de coordinación general, inducción de VPN, cuentas de cobro y formato GC-M2-F24).
${customRulesPrompt}
LISTADO DE PRODUCCIONES A CLASIFICAR:
${JSON.stringify(itemsPayload, null, 2)}

INSTRUCCIÓN:
Clasifica cada elemento por su índice (0 a ${items.length - 1}) asignándole obligationId (1 a 7) y una breve justificación en español.
Asegúrate de que las 7 obligaciones tengan documentos asignados (cobertura 100%) y de priorizar la parte técnica en las obligaciones 1 a 6.`;

      const geminiCall = ai.models.generateContent({
        model: 'gemini-3.8-flash',
        contents: prompt,
        config: {
          responseMimeType: 'application/json',
          responseSchema: {
            type: Type.OBJECT,
            properties: {
              classifications: {
                type: Type.ARRAY,
                items: {
                  type: Type.OBJECT,
                  properties: {
                    index: { type: Type.INTEGER },
                    obligationId: { type: Type.INTEGER },
                    reason: { type: Type.STRING },
                  },
                  required: ['index', 'obligationId', 'reason'],
                },
              },
            },
            required: ['classifications'],
          },
        },
      });

      const timeoutPromise = new Promise((_, reject) =>
        setTimeout(() => reject(new Error('Gemini API timeout')), 8000)
      );

      const response = await Promise.race([geminiCall, timeoutPromise]);

      const parsed = JSON.parse(response.text.trim());
      if (parsed && Array.isArray(parsed.classifications) && parsed.classifications.length > 0) {
        // Map back to guarantee all items are answered
        const resultMap = new Map();
        parsed.classifications.forEach((c) => {
          const parsedId = parseInt(c.obligationId, 10);
          resultMap.set(c.index, {
            obligationId: Math.min(Math.max(isNaN(parsedId) ? 2 : parsedId, 1), 7),
            reason: c.reason || 'Clasificado con IA Gemini',
          });
        });

        let classifications = items.map((it, idx) => {
          if (resultMap.has(idx)) {
            return resultMap.get(idx);
          }
          return classifyByRules(it.name, it.folderContext, customRules);
        });

        // Balanceamos para garantizar cobertura total de las 7 obligaciones
        classifications = balanceObligationsCoverage(items, classifications);

        return res.json({
          provider: 'gemini',
          model: 'gemini-3.8-flash',
          classifications,
        });
      }
    } catch (err) {
      console.warn('[Gemini classification failed, using rules fallback]', err.message);
    }
  }

  // Deterministic fallback with balancing
  let classifications = items.map((it) => classifyByRules(it.name, it.folderContext, customRules));
  classifications = balanceObligationsCoverage(items, classifications);

  return res.json({
    provider: 'rules',
    classifications,
  });
});

// App fallback routing
app.get('*', (req, res) => {
  res.sendFile(path.join(__dirname, 'index.html'));
});

app.listen(PORT, '0.0.0.0', () => {
  console.log(`Server running on http://0.0.0.0:${PORT}`);
});

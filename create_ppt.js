const pptxgen = require('pptxgenjs');
const {
  imageSizingCrop,
  imageSizingContain,
  safeOuterShadow,
  warnIfSlideHasOverlaps,
  warnIfSlideElementsOutOfBounds,
} = require('/home/oai/share/slides/pptxgenjs_helpers');

// Assets (created/downloaded locally)
const DIR = '/home/oai/networking_project';
const PATH_LOGO = `${DIR}/logo_conectatech.png`;
const IMG_BOGOTA_COLOR = `${DIR}/bogota_color.jpg`;
const IMG_BOGOTA_AERIAL = `${DIR}/bogota_aerial.jpg`;
const IMG_NETWORKING_REFRESH = `${DIR}/img_networking_refreshments.jpg`;
const IMG_CONVENTION = `${DIR}/img_convention_crowd.jpg`;
const IMG_LAPTOP = `${DIR}/img_laptop_coding.jpg`;

// Theme
const pptx = new pptxgen();
pptx.layout = 'LAYOUT_WIDE';

const COLORS = {
  black: '111111',
  white: 'FFFFFF',
  gray1: 'F5F6F7',
  gray2: 'E7EAEE',
  gray3: '6B7280',
  green: '2ECC71',
  greenDark: '1D8F52',
};

function addTopBar(slide, title) {
  slide.addShape(pptx.ShapeType.rect, {
    x: 0,
    y: 0,
    w: 13.333,
    h: 0.55,
    fill: { color: COLORS.black },
    line: { color: COLORS.black },
  });
  slide.addImage({ path: PATH_LOGO, ...imageSizingContain(PATH_LOGO, 0.15, 0.08, 2.6, 0.39) });
  slide.addText(title, {
    x: 2.85,
    y: 0.12,
    w: 10.25,
    h: 0.32,
    fontFace: 'Calibri',
    fontSize: 16,
    color: COLORS.white,
    bold: true,
  });
}

function addFooter(slide, text) {
  slide.addText(text, {
    x: 0.6,
    y: 7.18,
    w: 12.2,
    h: 0.25,
    fontFace: 'Calibri',
    fontSize: 10,
    color: COLORS.gray3,
  });
}

// ----------------------------
// Slide 1: Cover
// ----------------------------
{
  const slide = pptx.addSlide();
  slide.addImage({ path: IMG_CONVENTION, ...imageSizingCrop(IMG_CONVENTION, 0, 0, 13.333, 7.5) });
  slide.addShape(pptx.ShapeType.rect, {
    x: 0,
    y: 0,
    w: 13.333,
    h: 7.5,
    fill: { color: '000000', transparency: 45 },
    line: { color: '000000', transparency: 100 },
  });
  slide.addImage({ path: PATH_LOGO, ...imageSizingContain(PATH_LOGO, 0.7, 0.65, 6.2, 1.5) });
  slide.addText('Propuesta de evento de networking', {
    x: 0.7,
    y: 2.25,
    w: 12.0,
    h: 0.6,
    fontFace: 'Calibri',
    fontSize: 36,
    color: COLORS.white,
    bold: true,
  });
  slide.addText('ConectaTech by Platzi | Bogotá + Híbrido', {
    x: 0.7,
    y: 2.95,
    w: 12.0,
    h: 0.35,
    fontFace: 'Calibri',
    fontSize: 18,
    color: COLORS.white,
  });
  slide.addShape(pptx.ShapeType.roundRect, {
    x: 0.7,
    y: 3.45,
    w: 4.55,
    h: 0.6,
    fill: { color: COLORS.green },
    line: { color: COLORS.green },
    radius: 12,
    shadow: safeOuterShadow('000000', 0.25, 45, 3, 1.5),
  });
  slide.addText('Semana 6 – Networking', {
    x: 0.95,
    y: 3.58,
    w: 4.05,
    h: 0.35,
    fontFace: 'Calibri',
    fontSize: 14,
    color: COLORS.black,
    bold: true,
  });
  addFooter(slide, 'Complete nombres, ficha, fecha y detalles finales antes de entregar.');

  slide.addNotes(`
[Sources]
- Imagen de fondo (crowd/convention) – Pexels: https://www.pexels.com/photo/crowded-convention-center-gathering-event-30324916/
[/Sources]
`);
}

// ----------------------------
// Slide 2: Contexto / Justificación
// ----------------------------
{
  const slide = pptx.addSlide();
  addTopBar(slide, 'Contexto y justificación');
  slide.addShape(pptx.ShapeType.rect, { x: 0, y: 0.55, w: 13.333, h: 6.95, fill: { color: COLORS.gray1 }, line: { color: COLORS.gray1 } });

  slide.addText('¿Por qué un evento de networking?', {
    x: 0.7,
    y: 1.05,
    w: 12,
    h: 0.5,
    fontFace: 'Calibri',
    fontSize: 26,
    bold: true,
    color: COLORS.black,
  });

  slide.addShape(pptx.ShapeType.roundRect, {
    x: 0.7,
    y: 1.65,
    w: 7.65,
    h: 3.35,
    fill: { color: COLORS.white },
    line: { color: COLORS.gray2 },
    radius: 12,
    shadow: safeOuterShadow('000000', 0.10, 45, 2, 1),
  });

  const bullets = [
    'Capital social: relaciones que facilitan acceso a información, apoyo y oportunidades.',
    'Lazos débiles: conectan con información nueva y expanden opciones profesionales.',
    'Redes más conectadas se asocian con mejoras en colaboración y desempeño organizacional.',
  ];
  slide.addText(bullets.map((t) => ({ text: t, options: { bullet: { indent: 18 }, hanging: 6 } })), {
    x: 1.05,
    y: 1.9,
    w: 7.2,
    h: 3.0,
    fontFace: 'Calibri',
    fontSize: 16,
    color: COLORS.black,
    paraSpaceAfter: 6,
  });

  slide.addShape(pptx.ShapeType.roundRect, {
    x: 8.55,
    y: 1.65,
    w: 4.08,
    h: 4.95,
    fill: { color: COLORS.white },
    line: { color: COLORS.gray2 },
    radius: 12,
    shadow: safeOuterShadow('000000', 0.10, 45, 2, 1),
  });
  slide.addImage({ path: IMG_NETWORKING_REFRESH, ...imageSizingCrop(IMG_NETWORKING_REFRESH, 8.65, 1.75, 3.88, 2.55) });
  slide.addText('Networking intencional + seguimiento', {
    x: 8.75,
    y: 4.45,
    w: 3.68,
    h: 0.35,
    fontFace: 'Calibri',
    fontSize: 14,
    bold: true,
    color: COLORS.black,
  });
  slide.addText('Diseño: matchmaking, speed networking y agenda por intereses.', {
    x: 8.75,
    y: 4.82,
    w: 3.68,
    h: 0.8,
    fontFace: 'Calibri',
    fontSize: 12,
    color: COLORS.gray3,
    valign: 'top',
  });
  addFooter(slide, 'La propuesta se apoya en literatura de redes y capital social.');

  slide.addNotes(`
[Sources]
- Adler & Kwon (2002) – capital social: https://doi.org/10.5465/AMR.2002.5922314
- Granovetter (1973) – lazos débiles: https://doi.org/10.1086/225469
- McKinsey (2022) – redes y desempeño: https://www.mckinsey.com/capabilities/people-and-organizational-performance/our-insights/network-effects-how-to-rebuild-social-capital-and-improve-corporate-performance
- Foto (networking/refreshments) – Pexels: https://www.pexels.com/photo/people-eating-while-standing-at-the-table-8761556/
[/Sources]
`);
}

// ----------------------------
// Slide 3: Empresa seleccionada
// ----------------------------
{
  const slide = pptx.addSlide();
  addTopBar(slide, 'Empresa seleccionada');
  slide.addImage({ path: IMG_LAPTOP, ...imageSizingCrop(IMG_LAPTOP, 0, 0.55, 13.333, 6.95) });
  slide.addShape(pptx.ShapeType.rect, { x: 0, y: 0.55, w: 13.333, h: 6.95, fill: { color: '000000', transparency: 55 }, line: { color: '000000', transparency: 100 } });

  slide.addText('Platzi', { x: 0.85, y: 1.05, w: 6.0, h: 0.8, fontFace: 'Calibri', fontSize: 44, bold: true, color: COLORS.white });
  slide.addShape(pptx.ShapeType.rect, { x: 0.85, y: 1.9, w: 0.85, h: 0.08, fill: { color: COLORS.green }, line: { color: COLORS.green } });
  slide.addText('Escuela de tecnología en línea (mercado hispanohablante)', { x: 0.85, y: 2.05, w: 9.0, h: 0.35, fontFace: 'Calibri', fontSize: 18, color: COLORS.white });

  const cards = [
    { k: 'Qué hace', v: 'Formación online en habilidades digitales: programación, producto, data, diseño y marketing.' },
    { k: 'Trayectoria', v: 'Inicia en 2011 (Mejorando.la) y adopta el nombre Platzi en 2014.' },
    { k: 'Propósito', v: 'Aportar herramientas en tecnología para que más personas mejoren su calidad de vida.' },
  ];
  const x = 0.85;
  let y = 2.65;
  for (const c of cards) {
    slide.addShape(pptx.ShapeType.roundRect, {
      x,
      y,
      w: 11.65,
      h: 1.15,
      fill: { color: COLORS.white, transparency: 6 },
      line: { color: 'FFFFFF', transparency: 100 },
      radius: 12,
    });
    // Keep clear separation between label and description (avoid overlaps)
    slide.addText(c.k, { x: x + 0.35, y: y + 0.18, w: 1.75, h: 0.35, fontFace: 'Calibri', fontSize: 16, bold: true, color: COLORS.black });
    slide.addText(c.v, { x: x + 2.10, y: y + 0.18, w: 9.40, h: 0.8, fontFace: 'Calibri', fontSize: 14, color: COLORS.black, valign: 'top' });
    y += 1.35;
  }

  addFooter(slide, 'Nota: el evento es un diseño académico (no comunicación oficial).');

  slide.addNotes(`
[Sources]
- Perfil general: https://es.wikipedia.org/wiki/Platzi
- Descripción (artículo): https://elpais.com/tecnologia/2016/05/12/actualidad/1463013754_689019.html
- Visión de impacto social (blog): https://platzi.com/blog/platzi-y-nuestra-vision-de-impacto-social/
- Imagen (laptop/coding) – Unsplash: https://unsplash.com/photos/person-working-on-a-laptop-coding-and-analyzing-data-Su1XYFlftXA
[/Sources]
`);
}

// ----------------------------
// Slide 4: Objetivos
// ----------------------------
{
  const slide = pptx.addSlide();
  addTopBar(slide, 'Objetivos del evento');
  slide.addShape(pptx.ShapeType.rect, { x: 0, y: 0.55, w: 13.333, h: 6.95, fill: { color: COLORS.white }, line: { color: COLORS.white } });
  slide.addImage({ path: IMG_BOGOTA_AERIAL, ...imageSizingCrop(IMG_BOGOTA_AERIAL, 8.0, 0.55, 5.333, 6.95) });
  slide.addShape(pptx.ShapeType.rect, { x: 8.0, y: 0.55, w: 5.333, h: 6.95, fill: { color: '000000', transparency: 35 }, line: { color: '000000', transparency: 100 } });
  slide.addText('Bogotá (presencial) + streaming (híbrido)', { x: 8.25, y: 6.95, w: 4.8, h: 0.3, fontFace: 'Calibri', fontSize: 11, color: COLORS.white });

  slide.addText('Objetivo general', { x: 0.85, y: 1.05, w: 6.9, h: 0.4, fontFace: 'Calibri', fontSize: 18, bold: true, color: COLORS.greenDark });
  slide.addText('Fortalecer la red de contactos conectando talento, empresas y aliados para oportunidades de colaboración y crecimiento.', {
    x: 0.85,
    y: 1.45,
    w: 6.9,
    h: 0.85,
    fontFace: 'Calibri',
    fontSize: 16,
    color: COLORS.black,
    valign: 'top',
  });

  slide.addText('Objetivos específicos', { x: 0.85, y: 2.45, w: 6.9, h: 0.4, fontFace: 'Calibri', fontSize: 18, bold: true, color: COLORS.greenDark });
  const items = ['Matchmaking empresa ↔ talento', 'Alianzas B2B (formación y patrocinios)', 'Posicionamiento como hub', 'Pipeline medible (leads, vacantes, proyectos)'];
  slide.addText(items.map((t) => ({ text: t, options: { bullet: { indent: 18 }, hanging: 6 } })), { x: 1.1, y: 2.85, w: 6.55, h: 2.0, fontFace: 'Calibri', fontSize: 16, color: COLORS.black, paraSpaceAfter: 6 });

  slide.addShape(pptx.ShapeType.roundRect, { x: 0.85, y: 5.35, w: 6.9, h: 1.35, fill: { color: COLORS.gray1 }, line: { color: COLORS.gray2 }, radius: 12 });
  slide.addText('KPI sugeridos', { x: 1.1, y: 5.5, w: 6.4, h: 0.3, fontFace: 'Calibri', fontSize: 14, bold: true, color: COLORS.black });
  slide.addText('≥ 300 asistentes • ≥ 400 reuniones agendadas • ≥ 25 empresas • ≥ 6 patrocinios', { x: 1.1, y: 5.85, w: 6.4, h: 0.8, fontFace: 'Calibri', fontSize: 12, color: COLORS.gray3, valign: 'top' });

  slide.addNotes(`
[Sources]
- Imagen Bogotá (aérea) – Unsplash: https://unsplash.com/photos/an-aerial-view-of-a-city-street-and-buildings-cwfuGUIb-zw
[/Sources]
`);
}

// ----------------------------
// Slide 5: Concepto y agenda
// ----------------------------
{
  const slide = pptx.addSlide();
  addTopBar(slide, 'Concepto y agenda (1 día)');
  slide.addShape(pptx.ShapeType.rect, { x: 0, y: 0.55, w: 13.333, h: 6.95, fill: { color: COLORS.gray1 }, line: { color: COLORS.gray1 } });
  // Leave a bottom band for the footer to avoid placing text directly on the photo
  slide.addImage({ path: IMG_BOGOTA_COLOR, ...imageSizingCrop(IMG_BOGOTA_COLOR, 0, 0.55, 5.7, 6.40) });
  slide.addShape(pptx.ShapeType.rect, { x: 0, y: 0.55, w: 5.7, h: 6.40, fill: { color: '000000', transparency: 40 }, line: { color: '000000', transparency: 100 } });
  slide.addText('ConectaTech', { x: 0.55, y: 1.15, w: 4.9, h: 0.6, fontFace: 'Calibri', fontSize: 34, bold: true, color: COLORS.white });
  slide.addText('Formato: Presencial + streaming + app de reuniones', { x: 0.55, y: 1.8, w: 5.05, h: 0.5, fontFace: 'Calibri', fontSize: 12, color: COLORS.white });

  slide.addShape(pptx.ShapeType.roundRect, { x: 6.05, y: 1.05, w: 6.95, h: 5.95, fill: { color: COLORS.white }, line: { color: COLORS.gray2 }, radius: 12, shadow: safeOuterShadow('000000', 0.08, 45, 2, 1) });
  slide.addText('Agenda sugerida', { x: 6.35, y: 1.25, w: 6.4, h: 0.4, fontFace: 'Calibri', fontSize: 18, bold: true, color: COLORS.black });

  const agenda = [
    ['08:00', 'Registro + match check-in'],
    ['09:00', 'Keynote: tendencias y empleabilidad tech'],
    ['10:00', 'Speed networking (Bloque 1)'],
    ['11:00', 'Zonas temáticas + stands (recruiting)'],
    ['13:00', 'Almuerzo + reuniones 1:1'],
    ['14:30', 'Talleres cortos (CV, portafolio, entrevistas)'],
    ['16:00', 'Speed networking (Bloque 2)'],
    ['17:00', 'Pitch & Partners (startups/aliados)'],
    ['18:00', 'Cierre + after networking'],
  ];
  let y = 1.75;
  for (const [time, item] of agenda) {
    slide.addShape(pptx.ShapeType.roundRect, { x: 6.35, y, w: 6.35, h: 0.48, fill: { color: COLORS.gray1 }, line: { color: COLORS.gray2 }, radius: 8 });
    slide.addText(time, { x: 6.5, y: y + 0.11, w: 0.9, h: 0.28, fontFace: 'Calibri', fontSize: 12, bold: true, color: COLORS.greenDark });
    slide.addText(item, { x: 7.4, y: y + 0.11, w: 5.15, h: 0.28, fontFace: 'Calibri', fontSize: 12, color: COLORS.black });
    y += 0.56;
  }
  addFooter(slide, 'La agenda se ajusta a disponibilidad de aliados, speakers y venue.');

  slide.addNotes(`
[Sources]
- Imagen Bogotá (barrios coloridos) – Unsplash: https://unsplash.com/photos/aerial-view-of-town-IjIzgLEkwxs
[/Sources]
`);
}

// ----------------------------
// Slide 6: Público objetivo
// ----------------------------
{
  const slide = pptx.addSlide();
  addTopBar(slide, 'Público objetivo');
  slide.addShape(pptx.ShapeType.rect, { x: 0, y: 0.55, w: 13.333, h: 6.95, fill: { color: COLORS.white }, line: { color: COLORS.white } });
  slide.addText('Segmentos clave (valor mutuo)', { x: 0.7, y: 1.05, w: 12, h: 0.5, fontFace: 'Calibri', fontSize: 26, bold: true, color: COLORS.black });

  const segCards = [
    { title: 'Empresas empleadoras', tag: 'HR • Talent • Líderes técnicos', pts: ['Contratación', 'Marca empleadora', 'Alianzas de formación'] },
    { title: 'Startups', tag: 'Founders • PM • COO', pts: ['Talento', 'Partnerships', 'Clientes tempranos'] },
    { title: 'Comunidad Platzi', tag: 'Dev • Data • UX • Growth', pts: ['Empleo', 'Mentoría', 'Proyectos'] },
    { title: 'Aliados del ecosistema', tag: 'Universidades • Comunidades • VCs', pts: ['Difusión', 'Programas conjuntos', 'Innovación abierta'] },
  ];
  const grid = { x0: 0.7, y0: 1.75, w: 12.0, h: 5.2, cols: 2, rows: 2, gap: 0.35 };
  const cardW = (grid.w - grid.gap) / grid.cols;
  const cardH = (grid.h - grid.gap) / grid.rows;

  for (let i = 0; i < segCards.length; i++) {
    const r = Math.floor(i / grid.cols);
    const c = i % grid.cols;
    const x = grid.x0 + c * (cardW + grid.gap);
    const y = grid.y0 + r * (cardH + grid.gap);
    slide.addShape(pptx.ShapeType.roundRect, { x, y, w: cardW, h: cardH, fill: { color: COLORS.gray1 }, line: { color: COLORS.gray2 }, radius: 12 });
    slide.addText(segCards[i].title, { x: x + 0.35, y: y + 0.25, w: cardW - 0.7, h: 0.35, fontFace: 'Calibri', fontSize: 18, bold: true, color: COLORS.black });
    slide.addText(segCards[i].tag, { x: x + 0.35, y: y + 0.65, w: cardW - 0.7, h: 0.3, fontFace: 'Calibri', fontSize: 12, color: COLORS.gray3 });
    slide.addText(segCards[i].pts.map((t) => ({ text: t, options: { bullet: { indent: 16 }, hanging: 6 } })), { x: x + 0.45, y: y + 1.05, w: cardW - 0.9, h: cardH - 1.2, fontFace: 'Calibri', fontSize: 14, color: COLORS.black, paraSpaceAfter: 4 });
  }
  addFooter(slide, 'Criterio: perfiles con decisión/influencia y participantes con portafolio o propuesta clara.');
}

// ----------------------------
// Slide 7: Valor agregado
// ----------------------------
{
  const slide = pptx.addSlide();
  addTopBar(slide, 'Valor agregado');
  slide.addShape(pptx.ShapeType.rect, { x: 0, y: 0.55, w: 13.333, h: 6.95, fill: { color: COLORS.gray1 }, line: { color: COLORS.gray1 } });

  slide.addText('Diferenciadores del evento', { x: 0.7, y: 1.05, w: 12, h: 0.5, fontFace: 'Calibri', fontSize: 26, bold: true, color: COLORS.black });
  slide.addImage({ path: IMG_NETWORKING_REFRESH, ...imageSizingCrop(IMG_NETWORKING_REFRESH, 0.7, 1.75, 5.8, 4.8) });
  slide.addShape(pptx.ShapeType.rect, { x: 0.7, y: 1.75, w: 5.8, h: 4.8, fill: { color: '000000', transparency: 60 }, line: { color: '000000', transparency: 100 } });
  slide.addText('Matchmaking + speed networking', { x: 1.05, y: 6.05, w: 5.1, h: 0.35, fontFace: 'Calibri', fontSize: 14, bold: true, color: COLORS.white });

  const diffs = [
    ['Matchmaking previo', 'Formulario + afinidad para sugerir reuniones 1:1.'],
    ['Zonas temáticas', 'Data/IA, Desarrollo, Producto/UX, Growth, Talento/HR Tech.'],
    ['Clínicas exprés', 'CV, LinkedIn, portafolio y simulación de entrevista.'],
    ['Tablero de oportunidades', 'Vacantes, retos técnicos y propuestas de colaboración.'],
    ['Post-evento (30 días)', 'Comunidad privada + seguimiento de reuniones y acuerdos.'],
  ];
  let y = 1.75;
  for (const [t, d] of diffs) {
    slide.addShape(pptx.ShapeType.roundRect, { x: 6.75, y, w: 5.88, h: 0.92, fill: { color: COLORS.white }, line: { color: COLORS.gray2 }, radius: 12 });
    slide.addShape(pptx.ShapeType.roundRect, { x: 6.9, y: y + 0.18, w: 0.22, h: 0.56, fill: { color: COLORS.green }, line: { color: COLORS.green }, radius: 6 });
    slide.addText(t, { x: 7.2, y: y + 0.18, w: 5.3, h: 0.28, fontFace: 'Calibri', fontSize: 14, bold: true, color: COLORS.black });
    slide.addText(d, { x: 7.2, y: y + 0.47, w: 5.3, h: 0.36, fontFace: 'Calibri', fontSize: 12, color: COLORS.gray3 });
    y += 1.03;
  }
  addFooter(slide, 'Diferenciador central: diseño intencional de interacciones + seguimiento.');

  slide.addNotes(`
[Sources]
- Foto (networking/refreshments) – Pexels: https://www.pexels.com/photo/people-eating-while-standing-at-the-table-8761556/
[/Sources]
`);
}

// ----------------------------
// Slide 8: Precios
// ----------------------------
{
  const slide = pptx.addSlide();
  addTopBar(slide, 'Estructura de precios');
  slide.addShape(pptx.ShapeType.rect, { x: 0, y: 0.55, w: 13.333, h: 6.95, fill: { color: COLORS.white }, line: { color: COLORS.white } });
  slide.addText('Entradas y paquetes (COP)', { x: 0.7, y: 1.05, w: 12, h: 0.5, fontFace: 'Calibri', fontSize: 26, bold: true, color: COLORS.black });

  const tiers = [
    { name: 'Comunidad', price: '99.000', sub: 'Preventa 79.000', pts: ['Acceso general', 'Speed networking', 'App del evento'] },
    { name: 'Profesional', price: '189.000', sub: 'Preventa 149.000', pts: ['Todo Comunidad', 'Mentoría grupal', 'Kit digital'] },
    { name: 'VIP', price: '349.000', sub: 'Preventa 299.000', pts: ['Todo Profesional', 'Lounge VIP', 'Clínica CV/portafolio'], highlight: true },
  ];

  const startX = 0.7;
  const startY = 1.75;
  const cardW = 3.92;
  const cardH = 4.85;
  for (let i = 0; i < tiers.length; i++) {
    const x = startX + i * (cardW + 0.35);
    const y = startY;
    slide.addShape(pptx.ShapeType.roundRect, {
      x,
      y,
      w: cardW,
      h: cardH,
      fill: { color: tiers[i].highlight ? COLORS.green : COLORS.gray1 },
      line: { color: tiers[i].highlight ? COLORS.green : COLORS.gray2 },
      radius: 14,
      shadow: safeOuterShadow('000000', 0.10, 45, 2, 1),
    });
    slide.addText(tiers[i].name, { x: x + 0.35, y: y + 0.35, w: cardW - 0.7, h: 0.35, fontFace: 'Calibri', fontSize: 20, bold: true, color: COLORS.black });
    slide.addText(tiers[i].price, { x: x + 0.35, y: y + 0.85, w: cardW - 0.7, h: 0.6, fontFace: 'Calibri', fontSize: 34, bold: true, color: COLORS.black });
    slide.addText(tiers[i].sub, { x: x + 0.35, y: y + 1.43, w: cardW - 0.7, h: 0.3, fontFace: 'Calibri', fontSize: 12, color: tiers[i].highlight ? COLORS.black : COLORS.gray3 });
    slide.addShape(pptx.ShapeType.rect, { x: x + 0.35, y: y + 1.82, w: cardW - 0.7, h: 0.02, fill: { color: tiers[i].highlight ? COLORS.black : COLORS.gray2 }, line: { color: tiers[i].highlight ? COLORS.black : COLORS.gray2 }, transparency: 20 });
    slide.addText(tiers[i].pts.map((t) => ({ text: t, options: { bullet: { indent: 16 }, hanging: 6 } })), { x: x + 0.45, y: y + 2.05, w: cardW - 0.9, h: 1.6, fontFace: 'Calibri', fontSize: 14, color: COLORS.black, paraSpaceAfter: 4 });
    slide.addShape(pptx.ShapeType.roundRect, { x: x + 0.35, y: y + 3.95, w: cardW - 0.7, h: 0.6, fill: { color: tiers[i].highlight ? COLORS.black : COLORS.green }, line: { color: tiers[i].highlight ? COLORS.black : COLORS.green }, radius: 10 });
    slide.addText('Inscripción', { x: x + 0.35, y: y + 4.08, w: cardW - 0.7, h: 0.35, fontFace: 'Calibri', fontSize: 14, bold: true, align: 'center', color: COLORS.white });
  }

  // Remove the filled bar shape (it triggers severe text-overlap warnings); use a simple divider instead.
  slide.addShape(pptx.ShapeType.rect, { x: 0.7, y: 6.72, w: 12.0, h: 0.03, fill: { color: COLORS.gray2 }, line: { color: COLORS.gray2 } });
  slide.addText('Paquetes corporativos', { x: 0.7, y: 6.82, w: 3.2, h: 0.3, fontFace: 'Calibri', fontSize: 12, bold: true, color: COLORS.black });
  slide.addText('Recruiter Pass (2): 2.300.000 | Patrocinio: 7.500.000 (cupos limitados)', { x: 3.9, y: 6.82, w: 8.8, h: 0.3, fontFace: 'Calibri', fontSize: 12, color: COLORS.gray3 });
  slide.addText('Nota: precios propuestos con fines académicos; ajustar según costos reales.', {
    x: 0.7,
    y: 7.12,
    w: 11.8,
    h: 0.22,
    fontFace: 'Calibri',
    fontSize: 10,
    color: COLORS.gray3,
  });
}

// ----------------------------
// Slide 9: Estrategia de promoción
// ----------------------------
{
  const slide = pptx.addSlide();
  addTopBar(slide, 'Promoción (8 semanas)');
  slide.addShape(pptx.ShapeType.rect, { x: 0, y: 0.55, w: 13.333, h: 6.95, fill: { color: COLORS.gray1 }, line: { color: COLORS.gray1 } });

  slide.addText('Estrategia multicanal (comunidad + LinkedIn)', { x: 0.7, y: 1.05, w: 12, h: 0.5, fontFace: 'Calibri', fontSize: 26, bold: true, color: COLORS.black });

  const phases = [
    ['Semana 8-6', 'Lanzamiento: landing, email base, anuncio en LinkedIn, preventa'],
    ['Semana 6-4', 'Contenidos: casos de éxito, entrevistas a mentores, alianzas con comunidades'],
    ['Semana 4-2', 'Activación B2B: outreach a empresas, webinars teaser, cierre de patrocinios'],
    ['Semana 2-0', 'Cuenta regresiva: pauta segmentada, remarketing, agenda final y referidos'],
  ];

  // Timeline cards
  let y = 1.75;
  for (const [p, d] of phases) {
    slide.addShape(pptx.ShapeType.roundRect, { x: 0.7, y, w: 12.0, h: 1.15, fill: { color: COLORS.white }, line: { color: COLORS.gray2 }, radius: 12, shadow: safeOuterShadow('000000', 0.06, 45, 2, 1) });
    slide.addShape(pptx.ShapeType.roundRect, { x: 0.95, y: y + 0.25, w: 1.25, h: 0.65, fill: { color: COLORS.green }, line: { color: COLORS.green }, radius: 10 });
    slide.addText(p, { x: 0.95, y: y + 0.4, w: 1.25, h: 0.35, fontFace: 'Calibri', fontSize: 12, bold: true, align: 'center', color: COLORS.black });
    slide.addText(d, { x: 2.35, y: y + 0.28, w: 10.1, h: 0.8, fontFace: 'Calibri', fontSize: 14, color: COLORS.black, valign: 'top' });
    y += 1.35;
  }

  slide.addShape(pptx.ShapeType.roundRect, { x: 0.7, y: 6.82, w: 12.0, h: 0.68, fill: { color: COLORS.white }, line: { color: COLORS.gray2 }, radius: 12 });
  slide.addText('Mensaje clave: conecta con decisores fuera de tu círculo, crea lazos útiles y da seguimiento post-evento.', { x: 0.9, y: 6.90, w: 11.6, h: 0.28, fontFace: 'Calibri', fontSize: 12, color: COLORS.gray3 });
  slide.addText('Medición: CTR/registro, conversiones, reuniones agendadas, leads B2B y NPS.', { x: 0.9, y: 7.18, w: 11.6, h: 0.26, fontFace: 'Calibri', fontSize: 11, color: COLORS.gray3 });
}

// ----------------------------
// Slide 10: Presupuesto + KPI (gráfico)
// ----------------------------
{
  const slide = pptx.addSlide();
  addTopBar(slide, 'Viabilidad (presupuesto + KPI)');
  slide.addShape(pptx.ShapeType.rect, { x: 0, y: 0.55, w: 13.333, h: 6.95, fill: { color: COLORS.white }, line: { color: COLORS.white } });

  slide.addText('Distribución referencial de costos', { x: 0.7, y: 1.05, w: 12, h: 0.5, fontFace: 'Calibri', fontSize: 26, bold: true, color: COLORS.black });

  const dataChart = [
    { name: 'Costos', labels: ['Venue + montaje', 'Catering', 'Producción AV', 'Marketing', 'Plataforma + staff'], values: [35, 25, 15, 15, 10] },
  ];
  slide.addChart(pptx.ChartType.pie, dataChart, {
    x: 0.9,
    y: 1.65,
    w: 6.3,
    h: 5.15,
    showLegend: true,
    legendPos: 'r',
    dataLabelPosition: 'bestFit',
    dataLabelFormatCode: '0%'
  });

  slide.addShape(pptx.ShapeType.roundRect, { x: 7.55, y: 1.65, w: 5.08, h: 5.15, fill: { color: COLORS.gray1 }, line: { color: COLORS.gray2 }, radius: 12 });
  slide.addText('KPI de seguimiento (30/60/90 días)', { x: 7.85, y: 1.85, w: 4.6, h: 0.35, fontFace: 'Calibri', fontSize: 16, bold: true, color: COLORS.black });
  const kpis = [
    'Reuniones agendadas vs. realizadas',
    'Leads B2B calificados y acuerdos',
    'Vacantes publicadas y entrevistas',
    'Satisfacción (NPS) y retención comunidad',
  ];
  slide.addText(kpis.map((t) => ({ text: t, options: { bullet: { indent: 18 }, hanging: 6 } })), { x: 7.95, y: 2.3, w: 4.65, h: 2.6, fontFace: 'Calibri', fontSize: 14, color: COLORS.black, paraSpaceAfter: 6 });
  slide.addText('El seguimiento es el “segundo acto” del networking: convierte contactos en resultados.', { x: 7.85, y: 5.15, w: 4.75, h: 0.9, fontFace: 'Calibri', fontSize: 12, color: COLORS.gray3, valign: 'top' });
  addFooter(slide, 'El presupuesto es una estimación para explicar la lógica de costos.');
}

// ----------------------------
// Slide 11: Cierre + referencias
// ----------------------------
{
  const slide = pptx.addSlide();
  addTopBar(slide, 'Cierre');
  slide.addShape(pptx.ShapeType.rect, { x: 0, y: 0.55, w: 13.333, h: 6.95, fill: { color: COLORS.gray1 }, line: { color: COLORS.gray1 } });
  slide.addImage({ path: IMG_CONVENTION, ...imageSizingCrop(IMG_CONVENTION, 0, 0.55, 13.333, 6.95) });
  slide.addShape(pptx.ShapeType.rect, { x: 0, y: 0.55, w: 13.333, h: 6.95, fill: { color: '000000', transparency: 65 }, line: { color: '000000', transparency: 100 } });

  slide.addText('Conclusión', { x: 0.9, y: 1.05, w: 6.8, h: 0.6, fontFace: 'Calibri', fontSize: 34, bold: true, color: COLORS.white });
  slide.addText(
    'ConectaTech by Platzi es viable porque combina un propósito claro (empleabilidad + alianzas) con un diseño intencional de interacciones y seguimiento post-evento.',
    { x: 0.9, y: 1.7, w: 6.9, h: 1.25, fontFace: 'Calibri', fontSize: 16, color: COLORS.white, valign: 'top' }
  );
  slide.addShape(pptx.ShapeType.roundRect, { x: 0.9, y: 3.15, w: 6.9, h: 1.4, fill: { color: COLORS.white, transparency: 8 }, line: { color: 'FFFFFF', transparency: 100 }, radius: 12 });
  slide.addText('Siguientes pasos', { x: 1.15, y: 3.32, w: 6.4, h: 0.3, fontFace: 'Calibri', fontSize: 14, bold: true, color: COLORS.black });
  slide.addText(
    ['Definir fecha y venue', 'Confirmar aliados y mentores', 'Abrir inscripciones (landing)', 'Configurar KPIs y seguimiento'].map((t) => ({ text: t, options: { bullet: { indent: 18 }, hanging: 6 } })),
    { x: 1.15, y: 3.65, w: 6.4, h: 0.8, fontFace: 'Calibri', fontSize: 12, color: COLORS.black, paraSpaceAfter: 4 }
  );

  slide.addShape(pptx.ShapeType.roundRect, { x: 8.05, y: 1.05, w: 5.03, h: 5.5, fill: { color: COLORS.white, transparency: 8 }, line: { color: 'FFFFFF', transparency: 100 }, radius: 12 });
  slide.addText('Referencias (APA – resumen)', { x: 8.35, y: 1.25, w: 4.5, h: 0.3, fontFace: 'Calibri', fontSize: 14, bold: true, color: COLORS.black });
  slide.addText(
    [
      'Adler & Kwon (2002). Social capital…',
      'Granovetter (1973). Weak ties…',
      'McKinsey (2022). Network effects…',
      'El País (2016). Platzi, formación…',
      'Wikipedia (2019). Platzi.',
      'Platzi Blog (s. f.). Visión de impacto.',
    ].join('\n'),
    { x: 8.35, y: 1.6, w: 4.5, h: 4.8, fontFace: 'Calibri', fontSize: 11, color: COLORS.black, valign: 'top' }
  );

  addFooter(slide, 'Fin de la presentación.');

  slide.addNotes(`
[Sources]
- Adler & Kwon (2002): https://doi.org/10.5465/AMR.2002.5922314
- Granovetter (1973): https://doi.org/10.1086/225469
- McKinsey (2022): https://www.mckinsey.com/capabilities/people-and-organizational-performance/our-insights/network-effects-how-to-rebuild-social-capital-and-improve-corporate-performance
- El País (2016): https://elpais.com/tecnologia/2016/05/12/actualidad/1463013754_689019.html
- Wikipedia (Platzi): https://es.wikipedia.org/wiki/Platzi
- Platzi Blog (impacto social): https://platzi.com/blog/platzi-y-nuestra-vision-de-impacto-social/
- Foto fondo (crowd/convention) – Pexels: https://www.pexels.com/photo/crowded-convention-center-gathering-event-30324916/
[/Sources]
`);
}

// ----------------------------
// Layout QA
// ----------------------------
for (const s of pptx._slides) {
  warnIfSlideHasOverlaps(s, pptx);
  warnIfSlideElementsOutOfBounds(s, pptx);
}

// Write output
pptx.writeFile({ fileName: `${DIR}/Presentacion_ConectaTech_Platzi.pptx` });

// Flat-style SVG illustrations (no text inside, Korean labels are added in PPT)
const sharp = require("sharp");

const NAVY = "#14273E", ICE = "#E8F1F8", ORANGE = "#F26B1D", YEL = "#FFC20E", RED = "#D7263D", GREEN = "#2E9E5B";

function snowflakes(w, h, n, seed = 7, op = 0.8) {
  let s = seed, out = "";
  const r = () => (s = (s * 9301 + 49297) % 233280) / 233280;
  for (let i = 0; i < n; i++) out += `<circle cx="${(r() * w).toFixed(0)}" cy="${(r() * h).toFixed(0)}" r="${(2 + r() * 5).toFixed(1)}" fill="#fff" opacity="${(0.4 + r() * op * 0.6).toFixed(2)}"/>`;
  return out;
}

// person: x,y = feet center; vest color; scale
function person(x, y, s = 1, vest = ORANGE, opts = {}) {
  const arm = opts.armUp
    ? `<line x1="${x + 10 * s}" y1="${y - 78 * s}" x2="${x + 30 * s}" y2="${y - 118 * s}" stroke="${vest}" stroke-width="${9 * s}" stroke-linecap="round"/>`
    : `<line x1="${x + 10 * s}" y1="${y - 78 * s}" x2="${x + 22 * s}" y2="${y - 45 * s}" stroke="${vest}" stroke-width="${9 * s}" stroke-linecap="round"/>`;
  const flag = opts.flag
    ? `<line x1="${x + 30 * s}" y1="${y - 118 * s}" x2="${x + 30 * s}" y2="${y - 160 * s}" stroke="#333" stroke-width="${3 * s}"/><rect x="${x + 30 * s}" y="${y - 160 * s}" width="${26 * s}" height="${18 * s}" fill="${RED}"/>`
    : "";
  return `<g>
  <line x1="${x - 6 * s}" y1="${y - 40 * s}" x2="${x - 10 * s}" y2="${y}" stroke="#2b3a4a" stroke-width="${10 * s}" stroke-linecap="round"/>
  <line x1="${x + 6 * s}" y1="${y - 40 * s}" x2="${x + 10 * s}" y2="${y}" stroke="#2b3a4a" stroke-width="${10 * s}" stroke-linecap="round"/>
  <rect x="${x - 14 * s}" y="${y - 88 * s}" width="${28 * s}" height="${50 * s}" rx="${8 * s}" fill="${vest}"/>
  <rect x="${x - 14 * s}" y="${y - 66 * s}" width="${28 * s}" height="${5 * s}" fill="#e8e8e8"/>
  <line x1="${x - 10 * s}" y1="${y - 78 * s}" x2="${x - 22 * s}" y2="${y - 45 * s}" stroke="${vest}" stroke-width="${9 * s}" stroke-linecap="round"/>
  ${arm}${flag}
  <circle cx="${x}" cy="${y - 102 * s}" r="${13 * s}" fill="#f1c9a5"/>
  <path d="M${x - 15 * s},${y - 104 * s} a${15 * s},${14 * s} 0 0 1 ${30 * s},0 z" fill="${YEL}"/>
  <rect x="${x - 18 * s}" y="${y - 106 * s}" width="${36 * s}" height="${4 * s}" rx="2" fill="${YEL}"/>
</g>`;
}

function badge(x, y, n, col = RED) {
  return `<circle cx="${x}" cy="${y}" r="24" fill="${col}" stroke="#fff" stroke-width="4"/><text x="${x}" y="${y + 9}" font-family="DejaVu Sans, Arial" font-weight="bold" font-size="26" fill="#fff" text-anchor="middle">${n}</text>`;
}

function cone(x, y, s = 1) {
  return `<path d="M${x - 14 * s},${y} L${x - 4 * s},${y - 40 * s} L${x + 4 * s},${y - 40 * s} L${x + 14 * s},${y} z" fill="${ORANGE}"/><rect x="${x - 9 * s}" y="${y - 24 * s}" width="${18 * s}" height="${6 * s}" fill="#fff"/><rect x="${x - 18 * s}" y="${y - 3 * s}" width="${36 * s}" height="${5 * s}" fill="#333"/>`;
}

// side-view dump truck with plow & spreader
function snowTruck(x, y, s = 1, body = ORANGE) {
  return `<g transform="translate(${x},${y}) scale(${s})">
  <rect x="0" y="-150" width="300" height="110" rx="6" fill="${body}"/>
  <rect x="0" y="-160" width="300" height="14" fill="#c4520f"/>
  <rect x="300" y="-175" width="110" height="135" rx="14" fill="#fff"/>
  <rect x="330" y="-160" width="70" height="55" rx="6" fill="#9fc3e0"/>
  <rect x="350" y="-195" width="40" height="18" rx="4" fill="${YEL}"/>
  <rect x="-30" y="-110" width="34" height="70" fill="#555"/>
  <circle cx="-13" cy="-30" r="8" fill="#ddd"/>
  <path d="M410,-80 L470,-110 L480,-30 L410,-30 z" fill="${YEL}" stroke="#444" stroke-width="4"/>
  <rect x="0" y="-45" width="410" height="16" fill="#333"/>
  <circle cx="70" cy="-20" r="34" fill="#222"/><circle cx="70" cy="-20" r="13" fill="#888"/>
  <circle cx="160" cy="-20" r="34" fill="#222"/><circle cx="160" cy="-20" r="13" fill="#888"/>
  <circle cx="345" cy="-20" r="34" fill="#222"/><circle cx="345" cy="-20" r="13" fill="#888"/>
</g>`;
}

const svgs = {};

// A. cover scene
svgs.cover = (() => {
  const w = 1600, h = 900;
  return `<svg xmlns="http://www.w3.org/2000/svg" width="${w}" height="${h}">
  <rect width="${w}" height="${h}" fill="${NAVY}"/>
  <circle cx="1350" cy="160" r="70" fill="#f4f1d8" opacity="0.9"/>
  <path d="M0,560 L260,380 L460,520 L700,330 L980,540 L1220,400 L1600,560 L1600,900 L0,900 z" fill="#1f3a58"/>
  <path d="M0,620 L1600,620 L1600,900 L0,900 z" fill="#2c4a6b"/>
  <path d="M0,640 C400,610 1200,610 1600,640 L1600,900 L0,900 z" fill="#f4f8fb"/>
  <path d="M0,700 L1600,700 L1600,860 L0,860 z" fill="#4a5866"/>
  <path d="M0,700 L1600,700 L1600,712 L0,712 z" fill="#eef3f7"/>
  <path d="M0,848 L1600,848 L1600,860 L0,860 z" fill="#eef3f7"/>
  ${[0, 1, 2, 3, 4, 5, 6, 7].map(i => `<rect x="${i * 220 + 20}" y="776" width="110" height="10" fill="${YEL}"/>`).join("")}
  ${snowTruck(560, 830, 1.15)}
  <path d="M520,800 C470,810 430,840 380,850" stroke="#fff" stroke-width="8" stroke-dasharray="4 14" fill="none" opacity="0.8"/>
  ${cone(250, 840, 1.1)}${cone(1300, 840, 1.1)}
  ${person(1420, 840, 1.0, ORANGE, { armUp: true, flag: true })}
  ${snowflakes(w, 640, 140)}
</svg>`;
})();

// B. crane loading scene with hazards
svgs.crane = (() => {
  const w = 1400, h = 760;
  return `<svg xmlns="http://www.w3.org/2000/svg" width="${w}" height="${h}">
  <rect width="${w}" height="${h}" fill="#eaf2f8"/>
  <rect y="620" width="${w}" height="140" fill="#cfd8df"/>
  <path d="M0,620 L${w},620" stroke="#fff" stroke-width="10"/>
  <ellipse cx="420" cy="700" rx="230" ry="16" fill="#fff" opacity="0.8"/>
  <ellipse cx="1040" cy="680" rx="160" ry="12" fill="#bfe3ff" opacity="0.9"/>
  <!-- storage: stacked tonbags -->
  <g>${[[60, 520], [170, 520], [115, 420]].map(([x, y]) => `<rect x="${x}" y="${y}" width="100" height="100" rx="14" fill="#fff" stroke="#9aa9b5" stroke-width="4"/><path d="M${x + 25},${y} L${x + 35},${y - 22} M${x + 75},${y} L${x + 65},${y - 22}" stroke="#5c88b8" stroke-width="6"/>`).join("")}</g>
  <!-- crane truck (left) -->
  <g>
    <rect x="330" y="470" width="120" height="110" rx="12" fill="#fff" stroke="#999" stroke-width="3"/>
    <rect x="345" y="485" width="80" height="45" rx="6" fill="#9fc3e0"/>
    <rect x="450" y="520" width="300" height="60" fill="#5b6b7a"/>
    <rect x="470" y="440" width="40" height="80" fill="${YEL}"/>
    <line x1="490" y1="450" x2="820" y2="170" stroke="${YEL}" stroke-width="30" stroke-linecap="round"/>
    <line x1="820" y1="170" x2="880" y2="190" stroke="${YEL}" stroke-width="20" stroke-linecap="round"/>
    <line x1="880" y1="195" x2="880" y2="330" stroke="#333" stroke-width="4"/>
    <path d="M870,330 q10,20 20,0" stroke="#333" stroke-width="6" fill="none"/>
    <line x1="845" y1="370" x2="880" y2="335" stroke="#333" stroke-width="3"/><line x1="915" y1="370" x2="880" y2="335" stroke="#333" stroke-width="3"/>
    <rect x="830" y="368" width="100" height="100" rx="14" fill="#fff" stroke="#9aa9b5" stroke-width="4"/>
    <rect x="455" y="580" width="16" height="40" fill="#333"/><rect x="440" y="616" width="46" height="10" fill="#333"/>
    <rect x="725" y="580" width="16" height="40" fill="#333"/><rect x="710" y="616" width="46" height="10" fill="#333"/>
    <circle cx="390" cy="590" r="32" fill="#222"/><circle cx="560" cy="590" r="32" fill="#222"/><circle cx="660" cy="590" r="32" fill="#222"/>
  </g>
  <!-- danger radius under load -->
  <ellipse cx="880" cy="640" rx="150" ry="30" fill="${RED}" opacity="0.18" stroke="${RED}" stroke-width="4" stroke-dasharray="14 10"/>
  <!-- receiving dump truck (right) -->
  <g>
    <rect x="960" y="500" width="300" height="90" fill="${ORANGE}"/>
    <rect x="960" y="490" width="300" height="14" fill="#c4520f"/>
    <rect x="1260" y="455" width="110" height="135" rx="12" fill="#fff" stroke="#999" stroke-width="3"/>
    <rect x="1285" y="470" width="70" height="50" rx="6" fill="#9fc3e0"/>
    <circle cx="1020" cy="600" r="32" fill="#222"/><circle cx="1110" cy="600" r="32" fill="#222"/><circle cx="1310" cy="600" r="32" fill="#222"/>
  </g>
  <!-- worker on truck bed edge (danger) -->
  ${person(1030, 490, 0.95, ORANGE, { armUp: true })}
  <!-- worker under load (danger) -->
  ${person(800, 640, 0.9, "#7a8c9c")}
  <!-- signalman (safe) -->
  ${person(240, 690, 0.95, GREEN, { armUp: true })}
  ${badge(1080, 360, 1)}${badge(760, 470, 2)}${badge(930, 300, 3)}${badge(600, 400, 4)}${badge(1200, 650, 5)}${badge(470, 690, 6)}
  ${snowflakes(w, 300, 40, 3, 0.4)}
</svg>`;
})();

// C. work zone top-down
svgs.zone = (() => {
  const w = 1600, h = 520;
  let marks = "";
  for (let x = 0; x < w; x += 90) marks += `<rect x="${x}" y="254" width="50" height="8" fill="#fff"/>`;
  let cones = "";
  for (let i = 0; i < 7; i++) { const x = 560 + i * 45, y = 330 - i * 11; cones += `<circle cx="${x}" cy="${y + 60}" r="11" fill="${ORANGE}" stroke="#fff" stroke-width="3"/>`; }
  for (let i = 0; i < 9; i++) cones += `<circle cx="${880 + i * 55}" cy="${314}" r="11" fill="${ORANGE}" stroke="#fff" stroke-width="3"/>`;
  for (let i = 0; i < 4; i++) cones += `<circle cx="${1380 + i * 40}" cy="${314 + i * 18}" r="11" fill="${ORANGE}" stroke="#fff" stroke-width="3"/>`;
  return `<svg xmlns="http://www.w3.org/2000/svg" width="${w}" height="${h}">
  <rect width="${w}" height="${h}" fill="#f3f7fa"/>
  <rect y="120" width="${w}" height="290" fill="#5d6b78"/>
  <rect y="120" width="${w}" height="8" fill="#fff"/><rect y="402" width="${w}" height="8" fill="#fff"/>
  ${marks}
  <!-- zone shading -->
  <rect x="40" y="96" width="480" height="12" fill="${YEL}"/>
  <rect x="540" y="96" width="320" height="12" fill="${ORANGE}"/>
  <rect x="880" y="96" width="480" height="12" fill="${RED}"/>
  <rect x="1380" y="96" width="200" height="12" fill="${GREEN}"/>
  <!-- warning signs -->
  ${[110, 370, 500].map(x => `<rect x="${x - 4}" y="470" width="8" height="40" fill="#555"/><path d="M${x},${420} l30,52 h-60 z" fill="${YEL}" stroke="${RED}" stroke-width="6"/>`).join("")}
  ${cones}
  <!-- buffer truck (top-down) -->
  <rect x="800" y="290" width="130" height="80" rx="8" fill="${ORANGE}"/><rect x="900" y="295" width="30" height="70" fill="#fff"/>
  <path d="M800,300 h-20 v60 h20" fill="${YEL}" stroke="#333" stroke-width="3"/>
  ${[0, 1, 2].map(i => `<path d="M${812 + i * 18},318 l12,12 l-12,12" stroke="#fff" stroke-width="5" fill="none"/>`).join("")}
  <!-- workers in work zone (top-down circles) -->
  ${[1040, 1140, 1240].map(x => `<circle cx="${x}" cy="360" r="20" fill="${ORANGE}" stroke="#fff" stroke-width="4"/><circle cx="${x}" cy="360" r="9" fill="${YEL}"/>`).join("")}
  <!-- signalman on shoulder -->
  <circle cx="240" cy="455" r="22" fill="${GREEN}" stroke="#fff" stroke-width="4"/><circle cx="240" cy="455" r="10" fill="${YEL}"/>
  <rect x="262" y="435" width="34" height="22" fill="${RED}"/>
  <!-- approaching car -->
  <rect x="20" y="300" width="120" height="66" rx="16" fill="#2f80ed"/><rect x="85" y="308" width="30" height="50" rx="6" fill="#bcd9f7"/>
  <path d="M150,333 h60" stroke="#fff" stroke-width="6" stroke-dasharray="10 8"/><path d="M210,321 l20,12 l-20,12 z" fill="#fff"/>
  <!-- traffic flow lane 2 -->
  <rect x="1180" y="160" width="120" height="66" rx="16" fill="#9aa"/><rect x="1245" y="168" width="30" height="50" rx="6" fill="#dde"/>
  ${badge(200, 50, 1, "#b38600")}${badge(700, 50, 2, ORANGE)}${badge(1120, 50, 3, RED)}${badge(1480, 50, 4, GREEN)}
</svg>`;
})();

// D. icy road skid
svgs.ice = (() => {
  const w = 900, h = 620;
  return `<svg xmlns="http://www.w3.org/2000/svg" width="${w}" height="${h}">
  <rect width="${w}" height="${h}" fill="#1c3550"/>
  ${snowflakes(w, 260, 60, 11)}
  <path d="M0,330 L900,330 L900,620 L0,620 z" fill="#46596b"/>
  <path d="M120,360 C300,350 520,400 820,380 L840,560 C560,600 280,560 60,580 z" fill="#bfe6ff" opacity="0.35"/>
  <path d="M140,420 C260,410 330,470 430,470 M160,480 C280,470 350,520 450,515" stroke="#222" stroke-width="10" fill="none" opacity="0.55"/>
  <g transform="rotate(-18 560 470)">
    <rect x="460" y="410" width="230" height="90" rx="26" fill="${RED}"/>
    <rect x="500" y="375" width="140" height="55" rx="18" fill="${RED}"/>
    <rect x="515" y="385" width="50" height="38" rx="6" fill="#cfe6f7"/><rect x="575" y="385" width="50" height="38" rx="6" fill="#cfe6f7"/>
    <circle cx="510" cy="500" r="28" fill="#222"/><circle cx="640" cy="500" r="28" fill="#222"/>
  </g>
  ${[[700, 380], [740, 420], [720, 455]].map(([x, y]) => `<path d="M${x},${y} l40,-10" stroke="#fff" stroke-width="6" stroke-linecap="round"/>`).join("")}
  <path d="M740,180 l70,120 h-140 z" fill="${YEL}" stroke="${RED}" stroke-width="10"/>
  <rect x="736" y="300" width="10" height="40" fill="#999"/>
  <path d="M722,280 q18,-40 36,0" stroke="#222" stroke-width="7" fill="none"/>
  <path d="M708,280 l64,0" stroke="#222" stroke-width="6"/>
</svg>`;
})();

// E. truck bed fall
svgs.fall = (() => {
  const w = 900, h = 620;
  return `<svg xmlns="http://www.w3.org/2000/svg" width="${w}" height="${h}">
  <rect width="${w}" height="${h}" fill="#eaf2f8"/>
  <rect y="520" width="${w}" height="100" fill="#cfd8df"/>
  <ellipse cx="320" cy="535" rx="260" ry="14" fill="#bfe6ff"/>
  <rect x="120" y="290" width="520" height="150" fill="${ORANGE}"/>
  <rect x="120" y="280" width="520" height="16" fill="#c4520f"/>
  <rect x="640" y="230" width="170" height="210" rx="16" fill="#fff" stroke="#999" stroke-width="3"/>
  <rect x="680" y="250" width="100" height="70" rx="8" fill="#9fc3e0"/>
  <rect x="120" y="440" width="690" height="22" fill="#333"/>
  <circle cx="220" cy="480" r="45" fill="#222"/><circle cx="360" cy="480" r="45" fill="#222"/><circle cx="720" cy="480" r="45" fill="#222"/>
  <!-- falling worker -->
  <g transform="rotate(-50 110 330)">${person(110, 330, 1.05, "#7a8c9c")}</g>
  <path d="M60,380 q-30,60 -10,120" stroke="${RED}" stroke-width="6" fill="none" stroke-dasharray="12 10"/>
  <!-- height marker -->
  <line x1="860" y1="290" x2="860" y2="520" stroke="${RED}" stroke-width="6"/>
  <path d="M845,305 l15,-20 l15,20 M845,505 l15,20 l15,-20" stroke="${RED}" stroke-width="6" fill="none"/>
  <!-- ladder (safe) -->
  <g opacity="0.95"><line x1="560" y1="300" x2="590" y2="520" stroke="${GREEN}" stroke-width="10"/><line x1="610" y1="300" x2="640" y2="520" stroke="${GREEN}" stroke-width="10"/>
  ${[0, 1, 2, 3, 4].map(i => `<line x1="${566 + i * 5.5}" y1="${340 + i * 40}" x2="${616 + i * 5.5}" y2="${340 + i * 40}" stroke="${GREEN}" stroke-width="8"/>`).join("")}</g>
</svg>`;
})();

async function renderAll(dir) {
  const out = {};
  for (const [k, v] of Object.entries(svgs)) {
    const p = `${dir}/${k}.png`;
    await sharp(Buffer.from(v), { density: 144 }).png().toFile(p);
    out[k] = p;
  }
  return out;
}
module.exports = { renderAll };
if (require.main === module) renderAll(__dirname + "/img").then(console.log);

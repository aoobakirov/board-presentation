// Render react-icons to PNG at 512px, tinted, for use in the deck.
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const fa = require("react-icons/fa");
const gi = require("react-icons/gi");
const fs = require("fs");
const path = require("path");

const OUT = path.join(__dirname, "assets", "icons");
fs.mkdirSync(OUT, { recursive: true });

// icon key -> [pack, name, hexColor]
const ICONS = {
  leaf:      [fa, "FaLeaf",          "2D9C5A"],
  flask:     [fa, "FaFlask",         "2D9C5A"],
  seedling:  [fa, "FaSeedling",      "2D9C5A"],
  industry:  [fa, "FaIndustry",      "2D9C5A"],
  wheat:     [gi, "GiWheat",         "2D9C5A"],
  corn:      [gi, "GiCorn",          "2D9C5A"],
  drop:      [fa, "FaTint",          "2D9C5A"],
  recycle:   [fa, "FaRecycle",       "2D9C5A"],
  chart:     [fa, "FaChartLine",     "2D9C5A"],
  truck:     [fa, "FaTruck",         "2D9C5A"],
  handshake: [fa, "FaHandshake",     "2D9C5A"],
  check:     [fa, "FaCheck",         "2D9C5A"],
  pin:       [fa, "FaMapMarkerAlt",  "2D9C5A"],
  flag:      [fa, "FaFlagCheckered", "2D9C5A"],
  atom:      [fa, "FaAtom",          "2D9C5A"],
  microscope:[fa, "FaMicroscope",    "2D9C5A"],
  globe:     [fa, "FaGlobeAsia",     "2D9C5A"],
  bolt:      [fa, "FaBolt",          "2D9C5A"],
  layer:     [fa, "FaLayerGroup",    "2D9C5A"],
  boxes:     [fa, "FaBoxes",         "2D9C5A"],
  balance:   [fa, "FaBalanceScale",  "2D9C5A"],
  coins:     [fa, "FaCoins",         "2D9C5A"],
  network:   [fa, "FaProjectDiagram","2D9C5A"],
  user:      [fa, "FaUserTie",       "2D9C5A"],
  warehouse: [fa, "FaWarehouse",     "2D9C5A"],
  ship:      [fa, "FaShip",          "2D9C5A"],
  star:      [fa, "FaStar",          "2D9C5A"],
  arrowup:   [fa, "FaArrowUp",       "2D9C5A"],
  award:     [fa, "FaAward",         "2D9C5A"],
};

// colors to render each icon in (we generate white + green versions)
const VARIANTS = { white: "FFFFFF", green: "2D9C5A", dark: "14342B", amber: "E0A526" };

async function main() {
  for (const [key, [pack, name]] of Object.entries(ICONS)) {
    const Comp = pack[name];
    if (!Comp) { console.log("MISSING", name); continue; }
    for (const [vname, hex] of Object.entries(VARIANTS)) {
      const svg = ReactDOMServer.renderToStaticMarkup(
        React.createElement(Comp, { color: "#" + hex, size: 512 })
      );
      const buf = Buffer.from(svg);
      await sharp(buf, { density: 400 })
        .resize(512, 512, { fit: "contain", background: { r:0,g:0,b:0,alpha:0 } })
        .png()
        .toFile(path.join(OUT, `${key}_${vname}.png`));
    }
  }
  console.log("icons done");
}
main().catch(e => { console.error(e); process.exit(1); });

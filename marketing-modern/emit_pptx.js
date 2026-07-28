const pptxgen = require("pptxgenjs");
const path = require("path");
const { buildAll, C } = require("./build_lib.js");

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";
pres.defineSlideMaster({ title: "BASE", background: { color: C.light } });
buildAll(pres);
pres.writeFile({ fileName: path.join(__dirname, "..", "KazBioFert_Modern.pptx") })
  .then((f) => console.log("WROTE", f))
  .catch((e) => { console.error(e); process.exit(1); });

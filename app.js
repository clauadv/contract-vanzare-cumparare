const piz_zip = require("pizzip");
const docx_templater = require("docxtemplater");
const fs = require("fs");
const performance = require("perf_hooks").performance;

const content = fs.readFileSync("samples/contract.docx", "binary");
const zip = new piz_zip(content);
const doc = new docx_templater(zip, { paragraphLoop: true, linebreaks: true });

const data = JSON.parse(fs.readFileSync('data.json'));
doc.render(data);

const buffer = doc.getZip().generate({ type: "nodebuffer" });
fs.writeFile("bin/output.docx", buffer, (error) => {
    if (error) throw error;

    console.log(`\n[*] built in ${performance.now().toFixed(0)}ms`);
});
const piz_zip = require("pizzip");
const docx_templater = require("docxtemplater");
const fs = require("fs");
const performance = require("perf_hooks").performance;

const json = JSON.parse(fs.readFileSync('data.json'));
const type = eval(`json.type_${json.selected_type}`);

const content = fs.readFileSync(type.path, "binary");
const zip = new piz_zip(content);
const doc = new docx_templater(zip, { paragraphLoop: true, linebreaks: true });

doc.render(json.data);

const buffer = doc.getZip().generate({ type: "nodebuffer" });
fs.writeFile(`bin/output_type_${json.selected_type}.docx`, buffer, (error) => {
    if (error) throw error;

    console.log(`\n[*] built 'output_type_${json.selected_type}.docx' in ${performance.now().toFixed(0)}ms`);
});
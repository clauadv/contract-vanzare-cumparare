const piz_zip = require("pizzip");
const docx_templater = require("docxtemplater");
const fs = require("fs");
const performance = require("perf_hooks").performance;

const json = JSON.parse(fs.readFileSync('data.json'));
const get_type = () => {
    if (json.selected_type == 1)
        return { path: json.type_1.file };
    else if (json.selected_type == 2)
        return { path: json.type_2.file };

    return {};
}

const content = fs.readFileSync(get_type().path, "binary");
const zip = new piz_zip(content);
const doc = new docx_templater(zip, { paragraphLoop: true, linebreaks: true });

doc.render(json.data);

const buffer = doc.getZip().generate({ type: "nodebuffer" });
fs.writeFile(`bin/output_type_${json.selected_type}.docx`, buffer, (error) => {
    if (error) throw error;

    console.log(`\n[*] built 'output_type_${json.selected_type}.docx' in ${performance.now().toFixed(0)}ms`);
});
const express = require("express");
const excelJs = require("exceljs");
var fs = require("fs");

const app = express();

const PORT = 4000;

app.get("/export", async (req, res) => {
    try {
        let workbook = new excelJs.Workbook()

        const sheet = workbook.addWorksheet("relatorios")
        sheet.columns = [
            {header: "Agente", key:"agente", width: 50},
            {header: "Logradouro", key:"logradouro", width: 50},
            {header: "N° da residência", key:"numero", width: 100},
            {header: "Bairro", key:"bairro", width: 50},
            {header: "Nível da situação", key:"nivel", width: 50}
        ]

        let object = JSON.parse(fs.readFileSync('relatorios.json', 'utf8'))

        await object.relatorios.map((value, idx) => {
            sheet.addRow({
                agente: value.agente,
                logradouro: value.logradouro,
                numero: value.numero,
                bairro: value.bairro,
                nivel: value.nivel
            });
        });

        res.setHeader(
            "Content-Type",
            "application/vnd.openxmlformats-officedocument.spredsheetml.sheet"
        );
        res.setHeader(
            "Content-Disposition",
            "attachment;filename=" + "relatorios.xls"
        );

        workbook.xlsx.write(res)
    } catch (error) {
        console.log(error)
    }
})

app.listen(PORT, () => {
    console.log(`Server Running on PORT: ${PORT} `);
});
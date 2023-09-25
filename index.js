const express = require("express");
const excelJs = require("exceljs");
const jsonCases = require("./pagina"); //Pagina da varíavel com o Json dos casos
var fs = require("fs");

const app = express();

const PORT = 4000;

app.get("/export", async (req, res) => {
    try {
        let workbook = new excelJs.Workbook()

        const sheet = workbook.addWorksheet("relatorio")
        sheet.columns = [
            {header: "Latitude", key:"latitude", width: 50},
            {header: "Longitude", key:"longitude", width: 50},
            {header: "Status", key:"status", width: 50},
            {header: "ID do usuário", key:"user_id", width: 50},
            {header: "Data", key:"caseData", width: 50},
            {header: "Complemento", key:"complement", width: 50}
        ]

        await jsonCases.forEach((value, idx) => {
            sheet.addRow({
                latitude: value.latitude,
                longitude: value.longitude,
                status: value.status,
                user_id: value.user_id,
                caseData: value.caseData,
                complement: value.complement
            });
        });

        res.setHeader(
            "Content-Type",
            "application/vnd.openxmlformats-officedocument.spredsheetml.sheet"
        );
        res.setHeader(
            "Content-Disposition",
            "attachment;filename=" + "relatorio.xls"
        );

        workbook.xlsx.write(res)
    } catch (error) {
        console.log(error);
        res.status(500).json({ error: 'Erro ao exportar para Excel.' });
    }
})

app.listen(PORT, () => {
    console.log(`Server Running on PORT: ${PORT} `);
});

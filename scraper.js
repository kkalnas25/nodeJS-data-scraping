const axios = require("axios");
const cheerio = require("cheerio");
const Excel = require("exceljs");

const loadSite = async() => {

    const { data } = await axios.get("https://www.kwdowntown.com/agents/");

    let agents = [];
    const $ = cheerio.load(data);
    $(".agent-information").each((index, element) => {
        let agent = {
            name: $(element).children("h3").text(),
            tel: $(element).children("a").text().substring(0,12),
            email: ""
        }

        //handles case where email does not exist
        if($(element).children("a").next().prop("href")){
            agent.email = $(element).children("a").next().prop("href").substr(7);
        }

        agents.push(agent);
    });

    return agents;

}

loadSite().then(agents => {
    console.log(agents);
    let workbook = new Excel.Workbook();
    let worksheet = workbook.addWorksheet("Agents");
    worksheet.columns = [
        {header: "Name", key: "name"},
        {header: "Telephone", "key": "tel"},
        {header: "Email", "key": "email"}
    ];

    worksheet.columns.forEach(column => {
        column.width = column.header.length
    });
    worksheet.getRow(1).font = {bold: true}

    agents.forEach((e, index) => {
        const rowIndex = index + 2;

        worksheet.addRow({
            ...e
        })
    });

    workbook.xlsx.writeFile("Agents.xlsx");
});

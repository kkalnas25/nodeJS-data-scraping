const axios = require("axios");
const cheerio = require("cheerio");
const Excel = require("exceljs");

const loadSite = async() => {

    const { data } = await axios.get("https://www.kwdowntown.com/agents/");

    let agents = [];
    const $ = cheerio.load(data);

    //JQuery-like pulling of elements from the DOM
    $(".agent-information").each((index, element) => { //get each agent-information card
        let agent = {
            name: $(element).children("h3").text(), 
            tel: "", //data here needs to be sanitized
            email: "" //some emails are messed up, so we first need to verify that these exist
        }

        //handles case where tel does not exist
        if($(element).children("a").text()){
            agent.tel = $(element).children("a").text().substring(0,12);
            agent.tel = agent.tel.replace(/\D/g,''); //https://stackoverflow.com/questions/1862130/strip-all-non-numeric-characters-from-string-in-javascript
        }

        //get rid of weird case
        if(agent.tel === "Email"){
            agent.tel = "";
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
    let workbook = new Excel.Workbook();
    let worksheet = workbook.addWorksheet("Agents");
    worksheet.columns = [
        {header: "Name", key: "name"}, //create headers for our data. There is probably a better way to do this with Object.keys()
        {header: "Telephone", "key": "tel"},
        {header: "Email", "key": "email"}
    ];

    //format headers
    worksheet.columns.forEach(column => {
        column.width = column.header.length < 12 ? 12 : column.header.length
    });
    worksheet.getRow(1).font = {bold: true};

    //add data to the rows of the sheet. As you can see this is pretty friendly with Javascript objects
    agents.forEach((e, index) => {
        const rowIndex = index + 2;

        worksheet.addRow({
            ...e //our destructured object!
        });
    });

    //save to file
    workbook.xlsx.writeFile("Agents.xlsx");
});

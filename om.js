const puppeteer = require("puppeteer");
const dataurl = require("./url.json");
const data = require("./data.json");
const excelJS = require("exceljs");
const replace = require('replace-in-file');

const host = `http://localhost:3000`;
const pathproject = `E:/monosus github/orientalmotor_2023/`

(async () => {
  console.time();
  const browser = await puppeteer.launch({ headless: true });
  const page = await browser.newPage();
  await page.setDefaultNavigationTimeout(0);
  await page.setViewport({ width: 2160, height: 1024 });

  let datacheck = [];

  for (const url of dataurl) {
    await page.goto(`${host}${url}`);

    console.log(`path: ` + url);
    let getlink = await page.$$eval(
      ".l-main section.l-sec a,.l-main section.l-sec-title a",
      (divs) =>
        divs.map((div) => ({
          link: div.getAttribute("href"),
          modal: !!div.getAttribute("data-modal"),
        }))
    );

    for (const { link, modal } of getlink) {
      let path = link;
      let anchorlink = "";
      if (link !== null && !link.match(/^\s+$/)) {
        anchorlink = link.match(/#.+$/) || "";
        path = link.replace(anchorlink, "")
        path = path === "" && anchorlink != "" ? "AnchorLinkSingle" : path
      }

      let check = await page.evaluate(
        (data, path, anchorlink) => {
          let check = [];
          for (const checkurl of data) {
            if (checkurl.old === path || checkurl.new === path)
              check = [
                ...check,
                { old: !path ? "Blank" : checkurl.id, new: checkurl.new, anchorlink },
              ];
          }
          return check;
        },
        data,
        path,
        anchorlink
      );

      check = check[0]

      datacheck = [
        ...datacheck,
        {
          url,
          path,
          old: check?.old || "-",
          new: check?.new || "N/A",
          modal,
          anchorlink,
        },
      ];

      datachecked = datacheck.filter(
        (f) => ((f.new !== f.path && f.path !== "AnchorLinkSingle") || !f.path) && !f.modal
      ).map(m=> ({...m,newanchor: m.new+m.anchorlink}));
    }
  }
  await browser.close();
  // console.log(datachecked);

  const workbook = new excelJS.Workbook();
  const worksheet = workbook.addWorksheet("Oriental Motor");
  const path = "./files";
  worksheet.columns = [
    { header: "Checked Page", key: "url", width: 70 },
    { header: "Found Problem with Links", key: "path", width: 70 },
    { header: "Matched with Current Site (ID)", key: "old", width: 35 },
    { header: "The URL should be…", key: "newanchor", width: 70 },
  ];

  let firsturl = [];
  datachecked.forEach((data) => {
    if (!firsturl.includes(data.url)) firsturl = [...firsturl, data.url];
    else data.url = "";

    data.path = data.path+data.anchorlink
    data.newanchor = data.new === "N/A" ? "-" : data.newanchor

    worksheet.addRow(data);
  });
  worksheet.getRow(1).eachCell((cell) => {
    cell.font = { bold: true };
  });
  worksheet.eachRow(function (row, rowNumber) {
    row.eachCell(function (cell, colNumber) {
      if (rowNumber !== 1){
        if (cell.value != "-" && colNumber === 3)
          row.getCell(colNumber).font = {
            color: { argb: "00FF0000" },
          };
        if (cell.value.match(/[一-龠]+|[ぁ-ゔ]+|[ァ-ヴー]+|[々〆〤ヶ]/) && colNumber === 4)
          row.getCell(colNumber).font = {
            color: { argb: "004411D9" },
          };
      }
    });
  });
  try {
    const dataxlsx = await workbook.xlsx.writeFile(`${path}/OMG.xlsx`);
  } catch (err) {
    console.log("Error");
  }
  console.timeLog();
  console.log("Checked")

  let urlpath = ""
  datachecked.forEach((data) => {
    urlpath = data.url || urlpath
    if(!data.newanchor.match(/[一-龠]+|[ぁ-ゔ]+|[ァ-ヴー]+|[々〆〤ヶ]/) && data.newanchor != "-"){
      const results = replace.sync({
        files: `${pathproject}tmp/ejs${urlpath}index.ejs`,
        from: new RegExp(`href="${data.path}"`),
        to: `href="${data.newanchor}"`,
      });
    }
  })
  console.log("Changed")
})();
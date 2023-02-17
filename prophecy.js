let randoms = [];
let forecasts = [];

Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        document.getElementById("none").onchange = radioChange;
        document.getElementById("input").onchange = radioChange;
        document.getElementById("output").onchange = radioChange;
        document.getElementById("distro").onclick = distro;
        document.getElementById("montecarlo").onclick = montecarlo;

        Excel.run(context => {
            context.workbook.onSelectionChanged.add(workbookChange)
            context.workbook.worksheets.load("items")
            return context.sync().then(function(){
              if (context.workbook.worksheets.items.filter(f => f.name == "prophecy").length > 0) {
                console.log("prophecy found!")
              }
              else {
                let prophecy = context.workbook.worksheets.add("prophecy")
                range1 = prophecy.getRange("A1:E1");
                range1.values = [["name", "cell", "value", "distribution", "parameters"]];
                range1.format.fill.color = "yellow";
                range2 = prophecy.getRange("I1:K1");
                range2.values = [["name", "cell", "value"]];
                range2.format.fill.color = "red"
                range1 = prophecy.getRange("E1:G1").merge();
              }
          });
        });
      }
    });

async function workbookChange(event) {
    await Excel.run(async (context) => {
      let cell = context.workbook.getActiveCell();
      cell.load("address");
      return context.sync().then(function() {
        let address = cell.address
        if (randoms.includes(address)) {
          document.getElementById('input').checked = true;
          document.getElementById('distro').disabled = false;
        } else if (forecasts.includes(address)) {
          document.getElementById('output').checked = true;
          document.getElementById('distro').disabled = true;
        } else {
          document.getElementById('none').checked = true;
          document.getElementById('distro').disabled = true;
        }
      });
    });
}

async function radioChange(event) {
  await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getActiveWorksheet();
    let cell = context.workbook.getActiveCell();
    let prophecy = context.workbook.worksheets.getItem("prophecy")
    cell.load("address");
    cell.load("values")
    cell.load("numberFormat")
    return context.sync().then(function() {
      let address = cell.address
      let idx = randoms.indexOf(address);
      let idx2 = forecasts.indexOf(address);
      if (document.getElementById('input').checked) {
          document.getElementById('distro').disabled = false;
          if (idx == -1) {
            randoms.push(address)
            let row = randoms.length
            prophecy.getCell(row, 0).values = [["input_" + row]]
            prophecy.getCell(row, 1).hyperlink = {
                textToDisplay: address,
                screenTip: "input_" + row,
                documentReference: address
                }
            prophecy.getCell(row, 2).values = cell.values
            prophecy.getCell(row, 2).numberFormat = cell.numberFormat
            prophecy.getCell(row, 4).numberFormat = cell.numberFormat
            prophecy.getCell(row, 5).numberFormat = cell.numberFormat
            prophecy.getCell(row, 6).numberFormat = cell.numberFormat
            prophecy.getCell(row, 3).dataValidation.rule = {
                  list: {
                    inCellDropDown: true,
                    source: "uniform,normal,triangular"
                  }
                };
          }
          if (idx2 != -1) {
            forecasts.splice(idx2, 1);
            let range = prophecy.getRange("G" + (2+idx2) + ":I" + (2+idx2));
            range.delete(Excel.DeleteShiftDirection.up);
          }
          cell.format.fill.color = "yellow"
      } else if (document.getElementById('output').checked) {
          document.getElementById('distro').disabled = true;
          if (idx != -1) {
            randoms.splice(idx, 1);
            let range = prophecy.getRange("A" + (2+idx) + ":E" + (2+idx));
            range.delete(Excel.DeleteShiftDirection.up);
          }
          if (idx2 == -1) {
            forecasts.push(address);
            let row = forecasts.length
            prophecy.getCell(row, 6).values = [["output_" + row]]
            prophecy.getCell(row, 7).hyperlink = {
                textToDisplay: address,
                screenTip: "output_" + row,
                documentReference: address
                }
            prophecy.getCell(row, 8).values = cell.values
          }
          cell.format.fill.color = "red"
      } else {
          document.getElementById('distro').disabled = true;
          if (idx != -1) {
            randoms.splice(idx, 1);
            let range = prophecy.getRange("A" + (2+idx) + ":E" + (2+idx));
            range.delete(Excel.DeleteShiftDirection.up);
          }
          if (idx2 != -1) {
            forecasts.splice(idx2, 1);
            let range = prophecy.getRange("G" + (2+idx2) + ":I" + (2+idx2));
            range.delete(Excel.DeleteShiftDirection.up);
          }
          cell.format.fill.clear();
      }
    });
  });
}

async function distro(event) {
  await Excel.run(async (context) => {
    let cell = context.workbook.getActiveCell();
    cell.load("address");
    return context.sync().then(function() {
      row = 2 + randoms.indexOf(cell.address);
      let prophecy = context.workbook.worksheets.getItem("prophecy");
      prophecy.activate();
      let range = prophecy.getRange("A" + row + ":Z" + row);
      range.select()
    });
  });
}

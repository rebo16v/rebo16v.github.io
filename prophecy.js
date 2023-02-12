Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        document.getElementById("distro").onclick = distro;
        document.getElementById("none").onchange = radioChange;
        document.getElementById("input").onchange = radioChange;
        document.getElementById("output").onchange = radioChange;

        Excel.run(context => {
            context.workbook.onSelectionChanged.add(workbookChange)
            context.workbook.worksheets.load("items")
            return context.sync().then(function(){
              if (context.workbook.worksheets.items.filter(f => f.name == "prophecy").length > 0) {
                console.log("prophecy found!")
              }
              else {
                var prophecy = context.workbook.worksheets.add("prophecy")
                range = prophecy.getRange("A" + 1 + ":E" + 1);
                range.values = [["name", "cell", "value", "distribution", "parameters"]]
              }
          });
        });
      }
    });

async function workbookChange(event) {
    await Excel.run(async (context) => {
      var cell = context.workbook.getActiveCell();
      cell.load("address");
      return context.sync().then(function() {
        var address = cell.address
        console.log("workbookChange => " + address)
        if (randoms.includes(address)) {
          console.log("workbookChange => input")
          document.getElementById('input').checked = true;
          document.getElementById('distro').disabled = false;
        } else if (forecasts.includes(address)) {
          console.log("workbookChange => output")
          document.getElementById('output').checked = true;
          document.getElementById('distro').disabled = true;
        } else {
          console.log("workbookChange => none")
          document.getElementById('none').checked = true;
          document.getElementById('distro').disabled = true;
        }
      });
    });
}

async function radioChange(event) {
  console.log("radioChange => " + event.value);
  await Excel.run(async (context) => {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var cell = context.workbook.getActiveCell();
    var prophecy = context.workbook.worksheets.getItem("prophecy")
    cell.load("address");
    cell.load("values")
    return context.sync().then(function() {
      var address = cell.address
      var idx = randoms.indexOf(address)
      if (document.getElementById('input').checked) {
          document.getElementById('distro').disabled = false;
          if (idx == -1) {
            randoms.push(address)
            var row = randoms.length
            prophecy.getCell(row, 0).values = [["random_" + row]]
            // prophecy.getCell(row, 1).values = [[address]]
            prophecy.getCell(row, 1).hyperlink = {
                textToDisplay: address,
                screenTip: "random_" + row,
                documentReference: address
                }
            prophecy.getCell(row, 2).values = cell.values
            prophecy.getCell(row, 3).dataValidation.rule = {
                  list: {
                    inCellDropDown: true,
                    source: "uniform (min,max),normal (mean,stdev),triangular (mean,stdev,mode)"
                  }
                };
          }
          if (forecasts.indexOf(address) != -1) forecasts.splice(forecasts.indexOf(address), 1);
          cell.format.fill.color = "yellow"
      } else if (document.getElementById('output').checked) {
          document.getElementById('distro').disabled = true;
          if (idx != -1) {
            randoms.splice(idx, 1);
            var range = prophecy.getRange("A" + (2+idx) + ":Z" + (2+idx));
            range.delete(Excel.DeleteShiftDirection.up);
          }
          if (forecasts.indexOf(address) == -1) forecasts.push(address)
          cell.format.fill.color = "red"
      } else {
          document.getElementById('distro').disabled = true;
          if (idx != -1) {
            randoms.splice(idx, 1);
            var range = prophecy.getRange("A" + (2+idx) + ":Z" + (2+idx));
            range.delete(Excel.DeleteShiftDirection.up);
          }
          if (forecasts.indexOf(address) != -1) forecasts.splice(forecasts.indexOf(address), 1);
          cell.format.fill.clear();
      }
    });
  });
}

async function distro(event) {
  await Excel.run(async (context) => {
    var cell = context.workbook.getActiveCell();
    cell.load("address");
    return context.sync().then(function() {
      console.log("distro => " + cell.address);
      row = 2 + randoms.indexOf(cell.address);
      var prophecy = context.workbook.worksheets.getItem("prophecy");
      prophecy.activate();
      var range = prophecy.getRange("A" + row + ":Z" + row);
      range.select()
    });
  });
}

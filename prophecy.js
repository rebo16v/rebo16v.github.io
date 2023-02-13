var randoms = [];
var forecasts = [];

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

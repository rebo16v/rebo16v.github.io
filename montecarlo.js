async function montecarlo() {
  await Excel.run(async(context) => {
    const n_iter = parseInt(document.getElementById("niter").value);
    let app = context.workbook.application;
    app.suspendApiCalculationUntilNextSync();
    ranges.forEach(r => {
      let [s, c] = r.split("!");
      let s2 = context.workbook.worksheets.getItem(s);
      let c2 = s2.getRange(c);
      c2.values = [["input"]]
    });
    return context.sync().then(function() {
      forecasts.forEach(f => {
        let [s, c] = f.split("!");
        let s2 = context.workbook.worksheets.getItem(s);
        let c2 = s2.getRange(c);
        c2.load("values");
        context.sync().then(function() {
          console.log("f => " + c2.values);
        });
      });
    });
  });
}

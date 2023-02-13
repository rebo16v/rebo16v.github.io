async function montecarlo() {
  await Excel.run(async(context) => {
    const n_iter = parseInt(document.getElementById("niter").value);
    var prophecy = context.workbook.worksheets.getItem("prophecy");
    range = prophecy.getRange("A" + 2 + ":G" + (2+randoms.length));
    range.load("values");
    context.sync().then(function() {
      let conf = range.values;
      conf.forEach(c => {console.log("=> "+ c)}]);
      console.log("conf => " + conf)
      let app = context.workbook.application;
      app.suspendApiCalculationUntilNextSync();
      randoms.forEach((r,i) => {
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
  });
}

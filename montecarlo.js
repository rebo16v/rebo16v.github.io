let out = []

async function montecarlo() {
  await Excel.run(async(context) => {
    win = [];
    out = [];
    forecasts.forEach((f,i) => {
      out[i] = [];
      win[i] = window.open("simulation.html?id=" + i, "output-" + i);
    });
    const n_iter = parseInt(document.getElementById("niter").value);
    let app = context.workbook.application;
    var prophecy = context.workbook.worksheets.getItem("prophecy");
    range = prophecy.getRange("A" + 2 + ":G" + (1+randoms.length));
    range.load("values");
    await context.sync();
    let confs = range.values;
    for (let k = 0; k < n_iter; k++) {
      app.suspendApiCalculationUntilNextSync();
      console.log("iter => " + k);
      stepIn(confs, context);
      await context.sync()
      let outputs = stepOut(context);
      await context.sync()
      outputs.forEach((o,i) => {
        out[i].push(o.values);
        // win[i].document.title = "output-" + i;
        let element = win[i].document.getElementById("graph")
        element.style.backgroundColor = "#00FF00";
      });
    }
  });
  forecasts.forEach((f,i) => console.log(i + " => " + out[i].length));
}

function stepIn(confs, context) {
  confs.forEach(conf => {
    let input = 0;
    switch (conf[3]) {
      case "uniform":
        input = sampleUniform(conf[4], conf[5]);
      case "normal":
        input = sampleUniform(conf[4], conf[5]);
      case "triangular":
        input = sampleUniform(conf[4], conf[5], conf[6]);
    }
    let [s, c] = conf[1].split("!");
    let sheet = context.workbook.worksheets.getItem(s);
    let cell = sheet.getRange(c);
    cell.values = [[input]];
  });
}

function stepOut(context) {
  let ranges = [];
  forecasts.forEach(f => {
    let [s, c] = f.split("!");
    let sheet = context.workbook.worksheets.getItem(s);
    range = sheet.getRange(c);
    range.load("values");
    ranges.push(range);
  });
  return ranges;

}

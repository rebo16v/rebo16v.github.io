async function montecarlo() {
  await Excel.run(async(context) => {
    const n_iter = parseInt(document.getElementById("niter").value);
    let app = context.workbook.application;
    var prophecy = context.workbook.worksheets.getItem("prophecy");
    range = prophecy.getRange("A" + 2 + ":G" + (1+randoms.length));
    range.load("values");
    context.sync().then(function() {
      let confs = range.values;
      for (let i = 0; i < n_iter; i++) {
        app.suspendApiCalculationUntilNextSync();
        console.log("iter => " + i);
        stepIn(confs, context);
        context.sync().then(function() {
          stepOut(context);
        });
      }
    });
  });
}

async function stepIn(confs, context) {
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
  return context.sync();
}

async function stepOut(context) {
  let ranges = [];
  forecasts.forEach(f => {
    console.log("f => "+ f)
    let [s, c] = f.split("!");
    let sheet = context.workbook.worksheets.getItem(s);
    range = sheet.getRange(c);
    range.load("values");
    ranges.push(range);
  });
  context.sync().then(function() {
    ranges.forEach(r => {
      let output = r.values;
      console.log("output => " + output);
    });
  });
}

async function montecarlo() {
  await Excel.run(async(context) => {
    const n_iter = parseInt(document.getElementById("niter").value);
    let app = context.workbook.application;
    var prophecy = context.workbook.worksheets.getItem("prophecy");
    range = prophecy.getRange("A" + 2 + ":G" + (2+randoms.length));
    range.load("values");
    context.sync().then(function() {
      let confs = range.values;
      for (let i = 0; i < n_iter; i++) {
        app.suspendApiCalculationUntilNextSync();
        console.log("iter => " + i);
        stepIn(confs);
        context.sync().then(function() {
          let o = stepOut();
          console.log("output => " + o);
        });
      }
    });
  });
}




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

async function stepIn(confs) {
  confs.forEach(conf => {
    console.log("conf => "+ conf)
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

async function stepOut() {
  forecasts.forEach(f => {
    console.log("f => "+ f)
    let [s, c] = f.split("!");
    let sheet = context.workbook.worksheets.getItem(s);
    let cell = sheet.getRange(c);
    cell.load("values");
    context.sync().then(function() {
      let output = cell.values
      console.log("output => " + output);
    });
  });
}

let out = [];
let win = [];

async function montecarlo() {
  await Excel.run(async(context) => {
    win = [];
    out = [];
    forecasts.forEach((f,i) => {
      out[i] = [];
      /*
      Office.context.ui.displayDialogAsync("https://rebo16v.github.io/simulation.html",
          {height: 50, width: 50},
          function (asyncResult) {
              if (asyncResult.status == "failed") {
                console.log("error opening dialog => " +asyncResult.error.code);
              } else {
                  win[i] = asyncResult.value;
                  console.log("asyncResult");
              }
          });
          */
      win[i] = window.open("https://rebo16v.github.io/simulation.html");
    });
    console.log("hola!");
    await new Promise(r => setTimeout(r, 10000));
    console.log("adios!");
    // const n_iter = parseInt(document.getElementById("niter").value);
    let app = context.workbook.application;
    var prophecy = context.workbook.worksheets.getItem("prophecy");
    range = prophecy.getRange("A" + 2 + ":G" + (1+randoms.length));
    range.load("values");
    await context.sync();
    let confs = range.values;
    for (let k = 0; k < 5; k++) {
      app.suspendApiCalculationUntilNextSync();
      console.log("iter => " + k);
      stepIn(confs, context);
      await context.sync()
      let outputs = stepOut(context);
      await context.sync()
      outputs.forEach((o,i) => {
        let value = o.values[0][0]
        out[i].push(value);
        const msg = JSON.stringify({iter: k, value: value});
        win[i].postMessage(msg);
        console.log("message sent! " + msg);
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

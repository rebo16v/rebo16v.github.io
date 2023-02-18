let out;
let win;
let running;
let paused;

async function montecarlo() {
  await Excel.run(async(context) => {
    running = true;
    paused = false;
    let app = context.workbook.application;
    var prophecy = context.workbook.worksheets.getItem("prophecy");
    range_in = prophecy.getRange("A" + 2 + ":G" + (1+randoms.length));
    range_in.load("values");
    range_out = prophecy.getRange("I" + 2 + ":K" + (1+forecasts.length));
    range_out.load("values");
    await context.sync();
    let confs_in = range_in.values;
    let confs_out = range_out.values;
    win = [];
    out = [];
    let niter = parseInt(document.getElementById("niter").value);
    let nbins = parseInt(document.getElementById("nbins").value);
    confs_out.forEach((c,i) => {
      out[i] = [];
      win[i] = window.open("https://rebo16v.github.io/simulation.html?id=" + i + "&name=" + c[0] + "&nbins=" + nbins, "forecast_"+i);
    });
    await new Promise(r => setTimeout(r, 1000));
    for (let k = 0; k < niter; k++) {
      if (!running) break;
      while (paused) {await new Promise(r => setTimeout(r, 1000));}
      app.suspendApiCalculationUntilNextSync();
      stepIn(confs_in, context);
      await context.sync();
      let outputs = stepOut(confs_out, context);
      await context.sync();
      outputs.forEach((o,i) => {
        let value = o.values[0][0]
        out[i].push(value);
        // let msg = JSON.stringify({iter: k, value: value});
        win[i].postMessage(value);
      });
    }
  });
}

function stepIn(confs, context) {
  confs.forEach(conf => {
    let input = 0;
    switch (conf[3]) {
      case "uniform":
        input = sampleUniform(conf[4], conf[5]);
        break;
      case "normal":
        input = sampleNormal(conf[4], conf[5]);
        break;
      case "triangular":
        input = sampleTriangular(conf[4], conf[5], conf[6]);
        break;
    }
    let [s, c] = conf[1].split("!");
    let sheet = context.workbook.worksheets.getItem(s);
    let cell = sheet.getRange(c);
    cell.values = [[input]];
  });
}

function stepOut(confs, context) {
  let ranges = [];
  confs.forEach(conf => {
    let [s, c] = conf[1].split("!");
    let sheet = context.workbook.worksheets.getItem(s);
    range = sheet.getRange(c);
    range.load("values");
    ranges.push(range);
  });
  return ranges;
}

async function montecarlo_stop() {
  running = false;
}

async function montecarlo_start() {
  paused = false;
}

async function montecarlo_paused() {
  paused = true;
}

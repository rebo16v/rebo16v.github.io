const n_bins = 50

let randoms = []
let forecasts = []

const margin = {top: 50, right: 50, bottom: 50, left: 50},
    width = 500 - margin.left - margin.right,
    height = 300 - margin.top - margin.bottom;

var svg = d3.select("#montecarloGraph")
  .append("svg")
    .attr("width", width + margin.left + margin.right)
    .attr("height", height + margin.top + margin.bottom)
  .append("g")
    .attr("transform",
          "translate(" + margin.left + "," + margin.top + ")");

var axis = [svg.append("g")
  .attr("transform", "translate(0," + height + ")"),
  svg.append("g")]

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

  async function montecarlo() {
      await Excel.run(async(context) => {
          const sheet = context.workbook.worksheets.getActiveWorksheet();
          const n_iter = parseInt(document.getElementById("niter").value);
          const conf = sheet.getRange("B1:C2");
          conf.load("text");
          await context.sync();
          const [[revenues_mean, costs_mean], [revenues_stdev, costs_stdev]] = conf.text.map( x => x.map(z => parseInt(z)));
          console.log("revenues => (" + revenues_mean + "," + revenues_stdev + ")");
          console.log("costs => (" + costs_mean + "," + costs_stdev + ")");
          const revenues_exp = [revenues_mean - 3*revenues_stdev, revenues_mean + 3*revenues_stdev]
          const costs_exp = [costs_mean - 3*costs_stdev, costs_mean + 3*costs_stdev]

          var x = d3.scaleLinear()
              .domain([revenues_exp[0] - costs_exp[1], revenues_exp[1] - costs_exp[0]])
              .range([0, width]);
          axis[0].call(d3.axisBottom(x));

          var y = d3.scaleLinear()
                .domain([0, Math.round(5*n_iter/n_bins)])
                .range([height, 0]);
          axis[1].call(d3.axisLeft(y));

          const profits = [];
          profits_mean = 0;
          for (let i = 0; i < n_iter; i++) {
            const revenues = sampleNormal(revenues_mean, revenues_stdev);
            const costs = sampleNormal(costs_mean, costs_stdev);
            profits[i] = revenues - costs
            profits_mean = (profits_mean * (profits.length-1) + profits[i]) / profits.length
            sheet.getRange("D1").values = [[profits_mean]];
            sheet.getRange("A4:D4").values = [["iter-" + i, revenues, costs, profits[i]]];
            await context.sync();
            console.log("iter" + i + " => " + revenues + ", " + costs)

            var bins = d3.histogram()
                .domain(x.domain())
                .thresholds(x.ticks(50))
                (profits)

            svg.selectAll("rect")
                .data(bins)
                .join(
                    enter => enter
                        .append("rect")
                        .attr("x", function(d) {return x(d.x0)})
                        .attr("y", function(d) {return y(d.length)})
                        .attr("width", function(d) {return x(d.x1) - x(d.x0) - 2})
                        .attr("height", function(d) {return y(0) - y(d.length)})
                        .style("fill", "green"),
                    update => update
                        .attr("y", function(d) {return y(d.length)})
                        .attr("height", function(d) {return y(0) - y(d.length)}))

          }

          const profits_stdev = Math.sqrt(profits
            .map(k => (k - profits_mean)**2)
            .reduce((a,b) => a+b, 0 ) / profits.length);

          sheet.getRange("D2").values = [[profits_stdev]];
          await context.sync();
          return;

      });
  }

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

  function boxMullerTransform() {
  const u1 = Math.random();
  const u2 = Math.random();
  const z0 = Math.sqrt(-2.0 * Math.log(u1)) * Math.cos(2.0 * Math.PI * u2);
  const z1 = Math.sqrt(-2.0 * Math.log(u1)) * Math.sin(2.0 * Math.PI * u2);
  return { z0, z1 };
  }

  function sampleNormal(mean, stddev) {
      const { z0, _ } = boxMullerTransform();
      return z0 * stddev + mean;
  }

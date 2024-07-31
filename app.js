document.addEventListener('DOMContentLoaded', (event) => {
    console.log("DOM fully loaded and parsed");
    document.getElementById('analyzeButton').addEventListener('click', processFile);
});

let globalDataArray = []; // Store data globally for link generation

function processFile() {
    const fileInput = document.getElementById('fileInput');
    const file = fileInput.files[0];
    if (!file) {
        alert("Please select a file.");
        return;
    }

    const reader = new FileReader();

    reader.onload = function(e) {
        const data = e.target.result;
        const workbook = XLSX.read(data, { type: 'binary' });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        const json = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

        console.log("Loaded data:", json);

        processData(json);
    };

    reader.onerror = function(e) {
        console.error("File reading error:", e);
    };

    reader.readAsBinaryString(file);
}

function processData(data) {
    var dataArray = [];
    var xValues = [], yValues = [], zValues = [];
    for (var i = 1; i < data.length; i++) { // Skip header row
        var row = data[i];
        if (Array.isArray(row) && row.length >= 4) { // Ensure at least 4 columns
            var x = parseFloat(row[0]);
            var y = parseFloat(row[1]);
            var z = parseFloat(row[2]);
            var name = row[3];
            var color = row[4]; // Color column

            xValues.push(x);
            yValues.push(y);
            zValues.push(z);

            dataArray.push({
                'x': x,
                'y': y,
                'z': z,
                'name': name,
                'color': color
            });
        }
    }
    console.log("Processed data array:", dataArray);  // Log data for debugging
    globalDataArray = dataArray; // Store data globally

    // Automatically set default Min/Max values
    var xMin = Math.min(...xValues);
    var xMax = Math.max(...xValues);
    var yMin = Math.min(...yValues);
    var yMax = Math.max(...yValues);
    var zMin = Math.min(...zValues);
    var zMax = Math.max(...zValues);

    createChart(dataArray, xMin, xMax, yMin, yMax, zMin, zMax);
}

function createChart(dataArray, xMin, xMax, yMin, yMax, zMin, zMax) {
    console.log("Creating chart with data:", dataArray);

    var basePixel = parseInt(document.getElementById('basePixel').value);
    var widthRatio = parseFloat(document.getElementById('widthRatio').value);
    var heightRatio = parseFloat(document.getElementById('heightRatio').value);
    var depthRatio = parseFloat(document.getElementById('depthRatio').value);

    var colorBy = document.getElementById('colorBy').value;

    if (colorBy === 'columnE') {
        var colors = ['red', 'green', 'blue', 'yellow', 'purple'];
        dataArray.forEach((point, index) => {
            point.color = colors[index % colors.length];
        });
    } else {
        dataArray.forEach((point) => {
            point.color = `rgba(0, 105, 255, ${1 - (point.z / zMax)})`;
        });
    }

    var chart = Highcharts.chart('container', {
        chart: {
            renderTo: 'container',
            type: 'scatter3d',
            options3d: {
                enabled: true,
                alpha: 10,
                beta: 30,
                depth: basePixel * depthRatio,
                viewDistance: 25,
                fitToPlot: false,
                frame: {
                    bottom: { size: 1, color: 'rgba(0,0,0,0.02)' },
                    back: { size: 1, color: 'rgba(0,0,0,0.04)' },
                    side: { size: 1, color: 'rgba(0,0,0,0.06)' }
                }
            }
        },
        title: {
            text: '3D Scatter Plot'
        },
        subtitle: {
            text: 'Use the mouse to navigate around this 3D plot.'
        },
        plotOptions: {
            scatter: {
                width: basePixel * widthRatio,
                height: basePixel * heightRatio,
                depth: basePixel * depthRatio,
                dataLabels: {
                    enabled: true,
                    format: '{point.name}',
                    style: {
                        fontSize: '12px',
                        color: 'black',
                        textOutline: 'none'
                    }
                }
            }
        },
        xAxis: {
            min: xMin,
            max: xMax,
            title: {
                text: document.getElementById('xAxisLabel').value
            }
        },
        yAxis: {
            min: yMin,
            max: yMax,
            title: {
                text: document.getElementById('yAxisLabel').value
            }
        },
        zAxis: {
            min: zMin,
            max: zMax,
            title: {
                text: document.getElementById('zAxisLabel').value
            }
        },
        series: [{
            name: 'Data',
            data: dataArray,
            colorByPoint: true,
            dataLabels: {
                enabled: true,
                format: '{point.name}',
                style: {
                    fontSize: '12px',
                    color: 'black',
                    textOutline: 'none'
                }
            }
        }],
        tooltip: {
            headerFormat: '',
            pointFormat: '<b>{point.name}</b><br>X: {point.x}<br>Y: {point.y}<br>Z: {point.z}'
        }
    });
    console.log(chart); // Log chart object for debugging

    // Add 3D scatter plot mouse events
    (function (H) {
        function dragStart(eStart) {
            eStart = chart.pointer.normalize(eStart);

            var posX = eStart.chartX,
                posY = eStart.chartY,
                alpha = chart.options.chart.options3d.alpha,
                beta = chart.options.chart.options3d.beta,
                sensitivity = 5,  // Lower is more sensitive
                handlers = [];

            function drag(e) {
                // Get e.chartX and e.chartY
                e = chart.pointer.normalize(e);

               Sure, let's continue from where we left off with the updated `app.js` code to include the functionalities you need:

### Updated `app.js`

Here's the continuation of the updated JavaScript code with the necessary features:

```javascript
                chart.update({
                    chart: {
                        options3d: {
                            alpha: alpha + (e.chartY - posY) / sensitivity,
                            beta: beta + (posX - e.chartX) / sensitivity
                        }
                    }
                }, undefined, undefined, false);
            }

            function unbindAll() {
                handlers.forEach(function (unbind) {
                    if (unbind) {
                        unbind();
                    }
                });
                handlers.length = 0;
            }

            handlers.push(H.addEvent(document, 'mousemove', drag));
            handlers.push(H.addEvent(document, 'touchmove', drag));
            handlers.push(H.addEvent(document, 'mouseup', unbindAll));
            handlers.push(H.addEvent(document, 'touchend', unbindAll));
        }
        H.addEvent(chart.container, 'mousedown', dragStart);
        H.addEvent(chart.container, 'touchstart', dragStart);
    }(Highcharts));
}

// Initial chart setup with default data
var defaultData = [
    [5.6, 7.0, 4.73, 'United States of America'],
    [8.07, 9.06, 4.48, "People's Republic of China"],
    [9.66, 10.55, 4.06, 'Japan'],
    [7.02, 8.58, 5.17, 'Great Britain'],
    [10.1, 9.29, 3.26, 'ROC']
];

processData(defaultData);

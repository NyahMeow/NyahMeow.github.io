document.addEventListener('DOMContentLoaded', (event) => {
    console.log("DOM fully loaded and parsed");
    document.getElementById('analyzeButton').addEventListener('click', processFile);
    document.getElementById('generateLinkButton').addEventListener('click', generateLink);
    checkAndLoadData(); // Check and load data if available in URL
});

let globalDataArray = []; // Store data globally for link generation

const SHEET_ID = 'YOUR_GOOGLE_SHEET_ID'; // Replace with your Google Sheet ID
const API_KEY = 'YOUR_GOOGLE_API_KEY'; // Replace with your Google API Key

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
    for (var i = 1; i < data.length; i++) { // Skip header row
        var row = data[i];
        if (Array.isArray(row) && row.length >= 4) { // Ensure at least 4 columns
            var x = parseFloat(row[0]);
            var y = parseFloat(row[1]);
            var z = parseFloat(row[2]);
            var name = row[3];
            dataArray.push({
                'x': x,
                'y': y,
                'z': z,
                'name': name
            });
        }
    }
    console.log("Processed data array:", dataArray);  // Log data for debugging
    globalDataArray = dataArray; // Store data globally
    createChart(dataArray);
}

function createChart(dataArray) {
    console.log("Creating chart with data:", dataArray);
    var chart = Highcharts.chart('container', {
        chart: {
            renderTo: 'container',
            type: 'scatter3d',
            options3d: {
                enabled: true,
                alpha: 10,
                beta: 30,
                depth: 350,
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
                width: 10,
                height: 10,
                depth: 10,
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
            title: {
                text: 'X Axis'
            }
        },
        yAxis: {
            title: {
                text: 'Y Axis'
            }
        },
        zAxis: {
            title: {
                text: 'Z Axis'
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

// Generate link using Google Sheets
async function generateLink() {
    const dataString = JSON.stringify(globalDataArray);

    try {
        const response = await fetch(`https://sheets.googleapis.com/v4/spreadsheets/${SHEET_ID}/values/A1:append?valueInputOption=USER_ENTERED&key=${API_KEY}`, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({
                range: "A1",
                majorDimension: "ROWS",
                values: [[dataString]]
            })
        });

        const result = await response.json();
        const baseURL = window.location.href.split('?')[0];
        const link = `${baseURL}?sheetId=${SHEET_ID}`;
        document.getElementById('generatedLink').value = link;
    } catch (error) {
        console.error("Error generating link:", error);
        alert("Failed to generate link. Please try again.");
    }
}

// Parse URL parameters to load chart data if available
async function checkAndLoadData() {
    const urlParams = new URLSearchParams(window.location.search);
    const sheetId = urlParams.get('sheetId');
    if (sheetId) {
        try {
            const response = await fetch(`https://sheets.googleapis.com/v4/spreadsheets/${sheetId}/values/A1:A?key=${API_KEY}`);
            const result = await response.json();
            const dataString = result.values[0][0];
            const dataArray = JSON.parse(dataString);
            createChart(dataArray);
        } catch (error) {
            console.error("Error loading data from URL:", error);
            alert("Failed to load data from the link. Please check the link and try again.");
        }
    }
}

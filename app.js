document.addEventListener('DOMContentLoaded', function() {
    document.getElementById('analyzeButton').addEventListener('click', processFile);
});

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
        
        const dataArray = [];
        for (let i = 1; i < json.length; i++) {
            const row = json[i];
            if (Array.isArray(row) && row.length >= 4) {
                dataArray.push({
                    x: parseFloat(row[0]),
                    y: parseFloat(row[1]),
                    z: parseFloat(row[2]),
                    name: row[3]
                });
            }
        }

        console.log('Data Array:', dataArray);
        createChart(dataArray);
    };

    reader.onerror = function(e) {
        console.error("File reading error:", e);
    };

    reader.readAsBinaryString(file);
}

function createChart(dataArray) {
    Highcharts.chart('container', {
        chart: {
            type: 'scatter3d',
            options3d: {
                enabled: true,
                alpha: 10,
                beta: 30,
                depth: 350
            }
        },
        title: {
            text: '3D Scatter Plot'
        },
        plotOptions: {
            scatter: {
                dataLabels: {
                    enabled: true,
                    format: '{point.name}'
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
            data: dataArray
        }]
    });
}

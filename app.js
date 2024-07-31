document.addEventListener('DOMContentLoaded', (event) => {
    console.log("DOM fully loaded and parsed");
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

        // 読み込まれたデータをコンソールに出力
        console.log("Loaded data:", json);

        processData(json);
    };

    reader.onerror = function(e) {
        console.error("ファイルの読み込み中にエラーが発生しました:", e);
    };

    reader.readAsBinaryString(file);
}

function processData(data) {
    var dataArray = [];
    for (var i = 1; i < data.length; i++) { // ヘッダー行をスキップ
        var row = data[i];
        if (Array.isArray(row) && row.length >= 4) { // 少なくとも4列があることを確認
            var x = parseFloat(row[0]);
            var y = parseFloat(row[1]);
            var z = parseFloat(row[2]);
            var country = row[3];
            var color = 'rgba(0, 105, 255, ' + (1 - (z / 5)) + ')'; // zの範囲に応じた色の設定
            dataArray.push({
                'x': x,
                'y': y,
                'z': z,
                'name': country,
                'color': color
            });
        }
    }
    console.log(dataArray);  // デバッグ用にデータを確認
    createChart(dataArray);
}

function createChart(dataArray) {
    console.log("Creating chart with data:", dataArray); // デバッグ用ログ
    var chart = Highcharts.chart('container', {
        chart: {
            renderTo: 'container',
            type: 'scatter3d',
            options3d: {
                enabled: true,
                alpha: 10,
                beta: 30,
                depth: 250,
                viewDistance: 5,
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
            text: 'Highcharts 3D Scatter Chart'
        },
        plotOptions: {
            scatter: {
                width: 10,
                height: 10,
                depth: 10,
                dataLabels: {
                    enabled: true,
                    format: '{point.name}'
                }
            }
        },
        yAxis: {
            min: 0,
            max: 30,
            title: {
                text: 'Y-Axis'
            }
        },
        xAxis: {
            min: 0,
            max: 30,
            gridLineWidth: 1,
            title: {
                text: 'X-Axis'
            }
        },
        zAxis: {
            min: 0,
            max: 30,
            showFirstLabel: false,
            title: {
                text: 'Z-Axis'
            }
        },
        legend: {
            enabled: false
        },
        series: [{
            name: 'Data',
            colorByPoint: true,
            data: dataArray
        }],
        tooltip: {
            headerFormat: '',
            pointFormat: '<b>{point.name}</b><br>GDP per capita: {point.x}<br>Population: {point.y}<br>Medals: {point.z}'
        }
    });
    console.log(chart); // チャートオブジェクトの確認

    // 3D散布図のマウスイベントを追加
    (function (H) {
        function dragStart(eStart) {
            eStart = chart.pointer.normalize(eStart);

            var posX = eStart.chartX,
                posY = eStart.chartY,
                alpha = chart.options.chart.options3d.alpha,
                beta = chart.options.chart.options3d.beta,
                sensitivity = 5,  // lower is more sensitive
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

// 初期チャートの設定とデフォルトデータ
var defaultData = [
    [5.6, 7.0, 4.73, 'United States of America'],
    [8.07, 9.06, 4.48, "People's Republic of China"],
    [9.66, 10.55, 4.06, 'Japan'],
    [7.02, 8.58, 5.17, 'Great Britain'],
    [10.1, 9.29, 3.26, 'ROC']
    // 必要に応じて追加のデフォルトデータを追加
];

processData(defaultData);

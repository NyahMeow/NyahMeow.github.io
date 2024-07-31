let chart; // グローバル変数としてチャートを宣言
let originalData; // 元のデータを保存する変数
let customColors = ['#FF0000', '#00FF00', '#0000FF', '#FFFF00', '#FFA500']; // 初期色設定

document.addEventListener('DOMContentLoaded', function() {
    const analyzeButton = document.getElementById('analyzeButton');
    const resizeChartButton = document.getElementById('resizeChart');
    const resizeChartByRatioButton = document.getElementById('resizeChartByRatio');
    const updateAxisRangeButton = document.getElementById('updateAxisRange');
    const updateUnitsButton = document.getElementById('updateUnits');
    const updateViewDistanceButton = document.getElementById('updateViewDistance');
    const updateColoringButton = document.getElementById('updateColoring');
    const updateColorsButton = document.getElementById('updateColors');

    analyzeButton.addEventListener('click', processFile);
    resizeChartButton.addEventListener('click', resizeChart);
    resizeChartByRatioButton.addEventListener('click', resizeChartByRatio);
    updateAxisRangeButton.addEventListener('click', updateAxisRange);
    updateUnitsButton.addEventListener('click', updateAxisUnits);
    updateViewDistanceButton.addEventListener('click', updateViewDistance);
    updateColoringButton.addEventListener('click', updateColoring);
    updateColorsButton.addEventListener('click', updateColors);
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

        // ファイル名から拡張子を除いた名前を取得 (追加)
        const fileName = file.name.split('.').slice(0, -1).join('.');

        // X, Y, Z軸の名称を初期設定
        if (json.length > 0) {
            document.getElementById('xAxisUnit').value = json[0][0] || 'X-Axis';
            document.getElementById('yAxisUnit').value = json[0][1] || 'Y-Axis';
            document.getElementById('zAxisUnit').value = json[0][2] || 'Z-Axis';
        }

        // チャートタイトルをファイル名に設定
        document.getElementById('container').setAttribute('data-title', fileName);
        
        originalData = json;
        processData(json);
    };

    reader.onerror = function(e) {
        console.error("ファイルの読み込み中にエラーが発生しました:", e);
    };

    reader.readAsBinaryString(file);
}

function resizeChart() {
    const width = parseInt(document.getElementById('chartWidth').value, 10);
    const height = parseInt(document.getElementById('chartHeight').value, 10);
    const depth = parseInt(document.getElementById('chartDepth').value, 10); // z軸のサイズを取得

    if (chart) {
        chart.update({
            chart: {
                width: width,
                height: height,
                options3d: {
                    depth: depth // z軸のサイズを更新
                }
            }
        });
    } 
}

function resizeChartByRatio() {
    const baseSize = parseInt(document.getElementById('baseSize').value, 10);
    const widthRatio = parseFloat(document.getElementById('widthRatio').value);
    const heightRatio = parseFloat(document.getElementById('heightRatio').value);
    const depthRatio = parseFloat(document.getElementById('depthRatio').value);

    const width = Math.round(baseSize * widthRatio);
    const height = Math.round(baseSize * heightRatio);
    const depth = Math.round(baseSize * depthRatio);

    if (chart) {
        chart.update({
            chart: {
                width: width,
                height: height,
                options3d: {
                    depth: depth // z軸のサイズを更新
                }
            }
        });
    }
}

function updateAxisUnits() {
    if (chart) {
        chart.xAxis[0].setTitle({ text: document.getElementById('xAxisUnit').value });
        chart.yAxis[0].setTitle({ text: document.getElementById('yAxisUnit').value });
        chart.zAxis[0].setTitle({ text: document.getElementById('zAxisUnit').value });
    }
}

function updateAxisRange() {
    if (chart) {
        chart.xAxis[0].update({
            min: parseFloat(document.getElementById('xMin').value),
            max: parseFloat(document.getElementById('xMax').value)
        });
        chart.yAxis[0].update({
            min: parseFloat(document.getElementById('yMin').value),
            max: parseFloat(document.getElementById('yMax').value)
        });
        chart.zAxis[0].update({
            min: parseFloat(document.getElementById('zMin').value),
            max: parseFloat(document.getElementById('zMax').value)
        });
    }
}

function updateViewDistance() {
    const viewDistance = parseInt(document.getElementById('viewDistance').value, 10);

    if (chart) {
        chart.update({
            chart: {
                options3d: {
                    viewDistance: viewDistance
                }
            }
        });
    }
}

function updateColoring() {
    const colorColumn = document.getElementById('colorColumn').value;
    processData(originalData, colorColumn);
}

function updateColors() {
    for (let i = 0; i < 5; i++) {
        customColors[i] = document.getElementById(`color${i}`).value;
    }
    processData(originalData, document.getElementById('colorColumn').value);
}

function processData(data, colorColumn = 'z') {
    const dataArray = [];
    let xMin = Infinity, xMax = -Infinity;
    let yMin = Infinity, yMax = -Infinity;
    let zMin = Infinity, zMax = -Infinity;
    
    for (let i = 1; i < data.length; i++) { // ヘッダー行をスキップ
        const row = data[i];
        if (Array.isArray(row) && row.length >= 4) { // 少なくとも4列があることを確認
            const x = parseFloat(row[0]);
            const y = parseFloat(row[1]);
            const z = parseFloat(row[2]);
            const name = row[3];
            const colorValue = colorColumn === 'z' ? z : parseFloat(row[4]); // 色付けに使用する値

            xMin = Math.min(xMin, x);
            xMax = Math.max(xMax, x);
            yMin = Math.min(yMin, y);
            yMax = Math.max(yMax, y);
            zMin = Math.min(zMin, z);
            zMax = Math.max(zMax, z);

            const color = colorColumn === 'z'
                ? 'rgba(0, 105, 255, ' + (colorValue / 5) + ')' // zの範囲に応じた色の設定（大きい方が濃い）
                : getColorByValue(colorValue); // E列の値に基づいた色の設定

            dataArray.push({
                'x': x,
                'y': y,
                'z': z,
                'name': name,
                'color': color
            });
        }
    }

    // Update the axis range inputs with calculated values
    document.getElementById('xMin').value = xMin;
    document.getElementById('xMax').value = xMax;
    document.getElementById('yMin').value = yMin;
    document.getElementById('yMax').value = yMax;
    document.getElementById('zMin').value = zMin;
    document.getElementById('zMax').value = zMax;

    createChart(dataArray);
}

function getColorByValue(value) {
    return customColors[value % customColors.length]; // 値に応じてカスタムカラーを循環させる
}

function createChart(dataArray) {
    const chartTitle = document.getElementById('container').getAttribute('data-title');
    const xAxisUnit = document.getElementById('xAxisUnit').value;
    const yAxisUnit = document.getElementById('yAxisUnit').value;
    const zAxisUnit = document.getElementById('zAxisUnit').value;
    const viewDistance = parseInt(document.getElementById('viewDistance').value, 10);
    chart = Highcharts.chart('container', {
        chart: {
            type: 'scatter3d',
            margin: [80, 80, 80, 80], // マージンを調整
            width: 600, // 初期幅
            height: 550, // 初期高さ
            options3d: {
                enabled: true,
                alpha: 10,
                beta: 30,
                depth: 320,
                viewDistance: viewDistance,
                fitToPlot: false, // fitToPlotをfalseに設定
                frame: {
                    bottom: { size: 1, color: 'rgba(0,0,0,0.02)' },
                    back: { size: 1, color: 'rgba(0,0,0,0.04)' },
                    side: { size: 1, color: 'rgba(0,0,0,0.06)' }
                }
            }
        },
        title: {
            text: chartTitle || '3D Scatter Plot' // ファイル名をタイトルに設定
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
                    formatter: function () {
                        return this.point.name;
                    },
                    style: {
                        fontSize: '10px',
                        color: 'black'
                    }
                }
            }
        },
        xAxis: {
            min: parseFloat(document.getElementById('xMin').value),
            max: parseFloat(document.getElementById('xMax').value),
            gridLineWidth: 1,
            title: {
                text: xAxisUnit
            }
        },
        yAxis: {
            min: parseFloat(document.getElementById('yMin').value),
            max: parseFloat(document.getElementById('yMax').value),
            title: {
                text: yAxisUnit
            }
        },
        zAxis: {
            min: parseFloat(document.getElementById('zMin').value),
            max: parseFloat(document.getElementById('zMax').value),
            showFirstLabel: false,
            title: {
                text: zAxisUnit
            }
        },
        legend: {
            enabled: false
        },
        series: [{
            name: 'Data',
            data: dataArray,
            colorByPoint: true
            dataLabels: {
                enabled: true,
                format: '{point.name}',
                style: {
                    fontSize: '8px',
                    color: 'gray',
                    textOutline: 'none'
                }
            },
        　keys: ['x', 'y', 'z', 'name', 'color'] // データのキーを指定
        }],
        tooltip: {
            headerFormat: '',
              pointFormat: `<b>{point.name}</b><br>X: {point.x:.2f}<br>Y: {point.y:.2f}<br>Z: {point.z:.2f}`
        }
    });

    // 3D散布図のマウスイベントを追加
    enable3DNavigation(chart);
}

function enable3DNavigation(chart) {
    // Enable 3D scatter plot mouse events
    (function (H) {
        function dragStart(eStart) {
            eStart = chart.pointer.normalize(eStart);

            const posX = eStart.chartX;
            const posY = eStart.chartY;
            const alpha = chart.options.chart.options3d.alpha;
            const beta = chart.options.chart.options3d.beta;
            const sensitivity = 5;  // lower is more sensitive
            const handlers = [];

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

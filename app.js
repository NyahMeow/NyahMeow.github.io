let chart; // グローバル変数としてチャートを宣言
let originalData; // 元のデータを保存する変数
let customColors = ['#FF0000', '#00FF00', '#0000FF', '#FFFF00', '#FFA500']; // 初期色設定

document.addEventListener('DOMContentLoaded', function() {
    const urlParams = new URLSearchParams(window.location.search);
    const chartConfigKey = urlParams.get('chartConfigKey');
    const controls = document.getElementById('controls');
    if (chartConfigKey) {
        const storedConfig = localStorage.getItem(chartConfigKey);
        if (storedConfig) {
            try {
                const chartConfig = JSON.parse(storedConfig);
                chartConfig.chart.options3d.enabled = true; // Enable 3D
                chartConfig.chart.options3d.alpha = 10;
                chartConfig.chart.options3d.beta = 30;
                chartConfig.chart.options3d.depth = 350;
                chartConfig.chart.options3d.viewDistance = 25;
                chartConfig.chart.options3d.fitToPlot = false;
                chartConfig.chart.options3d.frame = {
                    bottom: { size: 1, color: 'rgba(0,0,0,0.02)' },
                    back: { size: 1, color: 'rgba(0,0,0,0.04)' },
                    side: { size: 1, color: 'rgba(0,0,0,0.06)' }
                };

                chart = Highcharts.chart('container', chartConfig);
                controls.classList.add('hidden'); // Hide controls

                // Enable 3D navigation for the chart
                enable3DNavigation(chart);
            } catch (e) {
                console.error('Error parsing stored chart configuration:', e);
            }
        } else {
            console.error('No stored chart configuration found for key:', chartConfigKey);
        }
    }
    
    const analyzeButton = document.getElementById('analyzeButton');
    const resizeChartButton = document.getElementById('resizeChart');
    const resizeChartByRatioButton = document.getElementById('resizeChartByRatio');
    const updateAxisRangeButton = document.getElementById('updateAxisRange');
    const updateUnitsButton = document.getElementById('updateUnits');
    const updateViewDistanceButton = document.getElementById('updateViewDistance');
    const updateColoringButton = document.getElementById('updateColoring');
    const updateColorsButton = document.getElementById('updateColors');
    const generateLinkButton = document.getElementById('generateLink');

    analyzeButton.addEventListener('click', processFile);
    resizeChartButton.addEventListener('click', resizeChart);
    resizeChartByRatioButton.addEventListener('click', resizeChartByRatio);
    updateAxisRangeButton.addEventListener('click', updateAxisRange);
    updateUnitsButton.addEventListener('click', updateAxisUnits);
    updateViewDistanceButton.addEventListener('click', updateViewDistance);
    updateColoringButton.addEventListener('click', updateColoring);
    updateColorsButton.addEventListener('click', updateColors);
    generateLinkButton.addEventListener('click', generateLink);
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

        // ファイル名から拡張子を除いた名前を取得 (追加)
        const fileName = file.name.split('.').slice(0, -1).join('.');

        // X, Y, Z軸の名称を初期設定 (追加)
        if (json.length > 0) {
            document.getElementById('xAxisUnit').value = json[0][0] || 'X-Axis';
            document.getElementById('yAxisUnit').value = json[0][1] || 'Y-Axis';
            document.getElementById('zAxisUnit').value = json[0][2] || 'Z-Axis';
        }

        // チャートタイトルをファイル名に設定 (追加)
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
    console.log("Resizing chart...");
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
        console.log("Chart resized to width:", width, "height:", height, "depth:", depth);
    } else {
        console.error("Chart not found.");
    }
}

function resizeChartByRatio() {
    console.log("Resizing chart by ratio...");
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
        console.log("Chart resized by ratio to width:", width, "height:", height, "depth:", depth);
    } else {
        console.error("Chart not found.");
    }
}

function updateAxisUnits() {
    console.log("Updating axis units...");
    if (chart) {
        chart.xAxis[0].setTitle({ text: document.getElementById('xAxisUnit').value });
        chart.yAxis[0].setTitle({ text: document.getElementById('yAxisUnit').value });
        chart.zAxis[0].setTitle({ text: document.getElementById('zAxisUnit').value });
        console.log("Axis units updated.");
    } else {
        console.error("Chart not found.");
    }
}

function updateAxisRange() {
    console.log("Updating axis range...");
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
        console.log("Axis range updated.");
    } else {
        console.error("Chart not found.");
    }
}

function updateViewDistance() {
    console.log("Updating view distance...");
    const viewDistance = parseInt(document.getElementById('viewDistance').value, 10);

    if (chart) {
        chart.update({
            chart: {
                options3d: {
                    viewDistance: viewDistance
                }
            }
        });
        console.log("View distance updated to:", viewDistance);
    } else {
        console.error("Chart not found.");
    }
}

function updateColoring() {
    console.log("Updating coloring...");
    const colorColumn = document.getElementById('colorColumn').value;
    processData(originalData, colorColumn);
}

function updateColors() {
    console.log("Updating custom colors...");
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

    console.log(dataArray);  // デバッグ用にデータを確認
    createChart(dataArray);
}

function getColorByValue(value) {
    return customColors[value % customColors.length]; // 値に応じてカスタムカラーを循環させる
}

function createChart(dataArray) {
    console.log("Creating chart with data:", dataArray); // デバッグ用ログ
    const chartTitle = document.getElementById('container').getAttribute('data-title'); // 追加
    const xAxisUnit = document.getElementById('xAxisUnit').value; // 追加
    const yAxisUnit = document.getElementById('yAxisUnit').value; // 追加
    const zAxisUnit = document.getElementById('zAxisUnit').value; // 追加
    const viewDistance = parseInt(document.getElementById('viewDistance').value, 10); // 追加
    chart = Highcharts.chart('container', {
        chart: {
            renderTo: 'container',
            type: 'scatter3d',
            margin: [80, 80, 80, 80], // マージンを調整
            width: 650, // 初期幅
            height: 600, // 初期高さ
            options3d: {
                enabled: true,
                alpha: 10,
                beta: 30,
                depth: 350,
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
                depth: 10
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
            colorByPoint: true,
            data: dataArray,
            keys: ['x', 'y', 'z', 'name', 'color'] // データのキーを指定
        }],
        tooltip: {
            headerFormat: '',
            pointFormat: `<b>{point.name}</b><br>${xAxisUnit}: {point.x:.3f}<br>${yAxisUnit}: {point.y:.3f}<br>${zAxisUnit}: {point.z:.3f}` // 小数点以下3桁まで表示
        }
    });
    console.log(chart); // チャートオブジェクトの確認

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

function generateLink() {
    if (!chart) {
        alert("Chart is not initialized.");
        return;
    }

    const chartConfig = chart.userOptions;
    const chartConfigStr = JSON.stringify(chartConfig);
    const chartConfigKey = 'chartConfig_' + Date.now();
    localStorage.setItem(chartConfigKey, chartConfigStr);

    const generatedLink = `${window.location.origin}${window.location.pathname}?chartConfigKey=${chartConfigKey}`;
    const linkContainer = document.getElementById('linkContainer');
    linkContainer.innerHTML = `<a href="${generatedLink}" target="_blank">Open Chart</a>`;
}

<div class="p-4 bg-white rounded-lg shadow-xl border-t-4 border-[#0077B6]">
    <script src="https://cdn.tailwindcss.com"></script>
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <style>
        .chart-container {
            position: relative;
            width: 100%;
            max-width: 650px;
            margin-left: auto;
            margin-right: auto;
            height: 420px;
            max-height: 500px;
            padding: 10px;
        }
    </style>
    
    <div class="text-center mb-6">
        <h3 class="text-xl font-bold text-[#0077B6]">ESG Materiality Assessment Matrix (Bubble Plot)</h3>
        <p class="text-sm text-gray-600 mt-1">Evaluating issues based on Stakeholder Importance (Y-axis) vs. Financial Impact (X-axis).</p>
    </div>

    <div class="chart-container">
        <canvas id="materialityMatrix"></canvas>
    </div>
    
    <script>
        const ctxMatrix = document.getElementById('materialityMatrix').getContext('2d');
        
        // Data points (Issues) - format: {x: Financial Impact (0-100), y: Stakeholder Concern (0-100), r: Bubble Size (Issue Urgency)}
        // 氣泡數量增加到 10 個
        const issueData = [
            // Quadrant 1: High Materiality (Strategic Focus)
            { label: 'Climate Transition Strategy', x: 90, y: 85, r: 15, color: '#0077B6' },
            { label: 'Supply Chain Human Rights', x: 65, y: 95, r: 12, color: '#00B4D8' },
            { label: 'AI Ethics & Data Privacy', x: 80, y: 70, r: 10, color: '#48CAE4' },
            { label: 'Talent Retention & DEI', x: 70, y: 80, r: 10, color: '#90E0EF' }, 

            // Quadrant 2: Financial Focus (High Impact, Low Stakeholder)
            { label: 'Cyber Security Governance', x: 95, y: 40, r: 14, color: '#0077B6' },
            { label: 'Sustainable Finance/Tax', x: 85, y: 35, r: 9, color: '#00B4D8' },
            
            // Quadrant 3: Stakeholder Focus (Low Impact, High Stakeholder)
            { label: 'Community Engagement', x: 20, y: 80, r: 7, color: '#90E0EF' },
            
            // Quadrant 4: Lower Materiality (Monitoring)
            { label: 'Waste & Circularity', x: 45, y: 55, r: 7, color: '#CAF0F8' },
            { label: 'Responsible Marketing', x: 30, y: 45, r: 5, color: '#CAF0F8' },
            { label: 'Product Packaging', x: 55, y: 20, r: 6, color: '#CAF0F8' } 
        ];

        const chartJsTooltip = {
            callbacks: {
                title: function(tooltipItems) {
                    const item = tooltipItems[0];
                    // 標籤長度檢查與換行處理
                    let label = item.chart.data.labels[item.dataIndex];
                    if (Array.isArray(label)) { return label.join(' '); } else { return label; }
                },
                label: function(context) {
                    const data = context.dataset.data[context.dataIndex];
                    return [
                        `Issue: ${data.label}`,
                        `Financial Impact: ${data.x}%`,
                        `Stakeholder Concern: ${data.y}%`,
                        `Urgency/Size: ${data.r}`
                    ];
                }
            }
        };

        new Chart(ctxMatrix, {
            type: 'bubble',
            data: {
                // 將長標籤進行換行處理
                labels: issueData.map(d => {
                    const label = d.label;
                    if (label.length > 16) {
                        const words = label.split(' ');
                        let currentLine = '';
                        const lines = [];
                        words.forEach(word => {
                            if ((currentLine + ' ' + word).length <= 16) {
                                currentLine += (currentLine ? ' ' : '') + word;
                            } else {
                                lines.push(currentLine);
                                currentLine = word;
                            }
                        });
                        if (currentLine) lines.push(currentLine);
                        return lines;
                    }
                    return label;
                }),
                datasets: [{
                    label: 'Material Issues',
                    data: issueData,
                    backgroundColor: issueData.map(d => d.color),
                    borderColor: issueData.map(d => d.color),
                    borderWidth: 1.5,
                    hoverBackgroundColor: 'rgba(255, 255, 255, 0.8)',
                    hoverBorderColor: issueData.map(d => d.color),
                    hoverBorderWidth: 3
                }]
            },
            options: {
                maintainAspectRatio: false,
                scales: {
                    x: {
                        type: 'linear',
                        position: 'bottom',
                        min: 0, max: 100,
                        title: { display: true, text: 'Financial Impact (Low → High)' },
                        grid: { drawBorder: false }
                    },
                    y: {
                        type: 'linear',
                        position: 'left',
                        min: 0, max: 100,
                        title: { display: true, text: 'Stakeholder Importance (Low → High)' },
                        grid: { drawBorder: false }
                    }
                },
                plugins: {
                    legend: { display: false },
                    tooltip: chartJsTooltip
                }
            }
        });
    </script>
</div>

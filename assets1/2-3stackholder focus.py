<div class="p-4 bg-white rounded-lg shadow-xl border-t-4 border-[#00B4D8]">
    <script src="https://cdn.tailwindcss.com"></script>
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <style>
        .chart-container {
            position: relative;
            width: 100%;
            max-width: 550px;
            margin-left: auto;
            margin-right: auto;
            height: 400px;
            max-height: 500px;
            padding: 10px;
        }
    </style>
    
    <div class="text-center mb-6">
        <h3 class="text-xl font-bold text-[#00B4D8]">Social: Stakeholder Focus Areas (Perception Score)</h3>
        <p class="text-sm text-gray-600 mt-1">Average perception score (0-100) from key stakeholder groups.</p>
    </div>

    <div class="chart-container">
        <canvas id="stakeholderRadar"></canvas>
    </div>
    
    <script>
        const ctxStakeholder = document.getElementById('stakeholderRadar').getContext('2d');
        const chartJsTooltip = {
            callbacks: {
                title: function(tooltipItems) {
                    const item = tooltipItems[0];
                    let label = item.chart.data.labels[item.dataIndex];
                    if (Array.isArray(label)) { return label.join(' '); } else { return label; }
                }
            }
        };

        new Chart(ctxStakeholder, {
            type: 'radar',
            data: {
                labels: ['Climate Action', 'Ethics & Compliance', 'Human Capital', 'Data Security', 'Product Safety'],
                datasets: [{
                    label: 'Investors',
                    data: [90, 80, 55, 75, 60],
                    fill: true,
                    backgroundColor: 'rgba(0, 119, 182, 0.3)',
                    borderColor: '#0077B6',
                    pointBackgroundColor: '#0077B6',
                }, {
                    label: 'Customers/Employees',
                    data: [65, 70, 95, 85, 90],
                    fill: true,
                    backgroundColor: 'rgba(72, 202, 228, 0.3)',
                    borderColor: '#48CAE4',
                    pointBackgroundColor: '#48CAE4',
                }]
            },
            options: {
                maintainAspectRatio: false,
                scales: {
                    r: {
                        angleLines: { display: true },
                        suggestedMin: 0,
                        suggestedMax: 100,
                        ticks: { backdropColor: 'rgba(255, 255, 255, 0.7)' }
                    }
                },
                plugins: {
                    legend: { position: 'bottom' },
                    tooltip: chartJsTooltip
                }
            }
        });
    </script>
</div>

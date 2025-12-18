<div class="p-4 bg-white rounded-lg shadow-xl border-t-4 border-[#48CAE4]">
    <script src="https://cdn.tailwindcss.com"></script>
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <style>
        .chart-container {
            position: relative;
            width: 100%;
            max-width: 600px;
            margin-left: auto;
            margin-right: auto;
            height: 400px;
            max-height: 500px;
            padding: 10px;
        }
    </style>
    
    <div class="text-center mb-6">
        <h3 class="text-xl font-bold text-[#48CAE4]">Strategic Material Issues Ranking (Prioritization Score)</h3>
        <p class="text-sm text-gray-600 mt-1">Ranking of the top 8 issues based on combined financial and impact materiality scores.</p>
    </div>

    <div class="chart-container">
        <canvas id="topIssuesRanking"></canvas>
    </div>
    
    <script>
        const ctxRanking = document.getElementById('topIssuesRanking').getContext('2d');
        
        const issues = [
            'Climate Transition Strategy', 'Talent Retention', 'AI Ethics & Governance', 
            'Supply Chain Human Rights', 'Water Stewardship', 'Data Privacy', 
            'Circular Economy', 'Board Diversity'
        ];
        
        const scores = [95, 88, 85, 80, 72, 68, 60, 55];
        
        const chartJsTooltip = {
            callbacks: {
                title: function(tooltipItems) {
                    const item = tooltipItems[0];
                    let label = item.chart.data.labels[item.dataIndex];
                    if (Array.isArray(label)) { return label.join(' '); } else { return label; }
                }
            }
        };

        new Chart(ctxRanking, {
            type: 'bar',
            data: {
                labels: issues,
                datasets: [{
                    label: 'Prioritization Score (0-100)',
                    data: scores,
                    backgroundColor: ['#0077B6', '#00B4D8', '#48CAE4', '#90E0EF', '#CAF0F8', '#0077B6', '#00B4D8', '#48CAE4'],
                }]
            },
            options: {
                indexAxis: 'y', // Horizontal Bar Chart
                maintainAspectRatio: false,
                scales: {
                    x: {
                        min: 0, max: 100,
                        title: { display: true, text: 'Score' }
                    },
                    y: {
                        grid: { display: false }
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

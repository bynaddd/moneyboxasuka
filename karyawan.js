document.addEventListener("DOMContentLoaded", () => {
    // 1. Set Tanggal Hari Ini (Indonesian Locale)
    const dateElement = document.getElementById('currentDate');
    const today = new Date();
    const options = { day: 'numeric', month: 'long', year: 'numeric' };
    dateElement.textContent = today.toLocaleDateString('id-ID', options);

    // 2. Inisialisasi Chart Saldo
    initBalanceChart();
});

function initBalanceChart() {
    const ctx = document.getElementById('balanceChart').getContext('2d');

    // Membuat Gradient untuk Area Bawah Grafik (Soft Blue)
    const gradient = ctx.createLinearGradient(0, 0, 0, 180);
    gradient.addColorStop(0, 'rgba(59, 130, 246, 0.25)'); // Opacity biru di atas
    gradient.addColorStop(1, 'rgba(59, 130, 246, 0.0)');  // Transparan di bawah

    // Data simulasi untuk 6 bulan terakhir
    const dataLabels = ['Jan', 'Feb', 'Mar', 'Apr', 'Mei', 'Jun'];
    const dataValues = [32.5, 34.0, 36.5, 39.0, 43.1, 45.2];

    new Chart(ctx, {
        type: 'line',
        data: {
            labels: dataLabels,
            datasets: [{
                label: 'Total Saldo',
                data: dataValues,
                borderColor: '#3b82f6', // Soft Blue
                backgroundColor: gradient,
                borderWidth: 3,
                pointBackgroundColor: '#ffffff',
                pointBorderColor: '#3b82f6',
                pointBorderWidth: 2,
                pointRadius: 4,
                pointHoverRadius: 6,
                fill: true,
                tension: 0.4 // Membuat garis sangat smooth/melengkung (0.0 = kaku, 0.4 = smooth)
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: {
                    display: false // Sembunyikan legend (box warna)
                },
                tooltip: {
                    backgroundColor: '#1e293b', // Dark tooltip
                    padding: 12,
                    titleFont: { family: 'Outfit', size: 13, weight: 'normal' },
                    bodyFont: { family: 'Outfit', size: 15, weight: 'bold' },
                    displayColors: false,
                    cornerRadius: 8,
                    callbacks: {
                        label: function (context) {
                            return 'Rp ' + context.parsed.y + ' Jt';
                        }
                    }
                }
            },
            scales: {
                x: {
                    grid: {
                        display: false, // Hilangkan garis vertikal
                        drawBorder: false
                    },
                    ticks: {
                        font: { family: 'Outfit', size: 12 },
                        color: '#94a3b8' // Text abu muda
                    }
                },
                y: {
                    grid: {
                        color: '#f1f5f9', // Garis horizontal sangat tipis/soft
                        drawBorder: false,
                        tickLength: 0
                    },
                    ticks: {
                        display: false, // Sembunyikan angka di Y-Axis agar clean
                        min: 30, // Agar grafik tidak mulai dari 0 dan lebih kelihatan pergerakannya
                        max: 50
                    }
                }
            },
            interaction: {
                intersect: false,
                mode: 'index',
            },
            animation: {
                duration: 1500, // Animasi load chart lebih lambat dan smooth
                easing: 'easeOutQuart'
            }
        }
    });
}

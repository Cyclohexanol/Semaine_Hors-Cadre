// === Semaine Hors-Cadre — Frontend Logic ===

document.addEventListener('DOMContentLoaded', function () {
    const form = document.getElementById('upload-form');
    const progressArea = document.getElementById('progress-area');
    const stepsContainer = document.getElementById('progress-steps');
    const progressBarFill = document.getElementById('progress-bar-fill');
    const elapsedEl = document.getElementById('elapsed');
    const resultsSection = document.getElementById('results-section');
    const uploadSection = document.getElementById('upload-section');
    const fileInput = document.getElementById('file');
    const fileLabel = document.getElementById('file-label');

    // --- File input feedback ---
    if (fileInput && fileLabel) {
        fileInput.addEventListener('change', function () {
            if (this.files.length > 0) {
                const file = this.files[0];
                const ext = file.name.split('.').pop().toLowerCase();
                if (ext !== 'xlsx') {
                    fileLabel.textContent = 'Format invalide — utilisez un fichier .xlsx';
                    fileLabel.className = 'form-text text-danger';
                    return;
                }
                fileLabel.textContent = file.name;
                fileLabel.className = 'form-text text-success';
            } else {
                fileLabel.textContent = '';
            }
        });
    }

    // --- Form submission via SSE ---
    if (form) {
        form.addEventListener('submit', function (e) {
            e.preventDefault();

            const submitBtn = form.querySelector('button[type=submit]');
            if (submitBtn) {
                submitBtn.disabled = true;
                submitBtn.innerHTML = '<span class="spinner-border spinner-border-sm" role="status" aria-hidden="true"></span> Optimisation en cours...';
            }

            // Show progress, hide upload card
            progressArea.classList.add('active');

            // Start elapsed timer
            const startTime = Date.now();
            const timer = setInterval(function () {
                const secs = Math.floor((Date.now() - startTime) / 1000);
                const mins = Math.floor(secs / 60);
                elapsedEl.textContent = mins > 0 ? mins + 'm ' + (secs % 60) + 's' : secs + 's';
            }, 1000);

            // Send file via fetch
            const formData = new FormData(form);

            fetch('/optimize', {
                method: 'POST',
                body: formData
            })
            .then(function (resp) { return resp.json(); })
            .then(function (data) {
                if (data.error) {
                    clearInterval(timer);
                    showError(data.error);
                    return;
                }

                // Open SSE connection
                const jobId = data.job_id;
                const source = new EventSource('/progress/' + jobId);
                const completedSteps = [];
                let currentStep = null;

                source.addEventListener('progress', function (e) {
                    const event = JSON.parse(e.data);
                    updateProgress(event.step, event.pct, completedSteps, currentStep);
                    currentStep = event.step;
                    completedSteps.push(event.step);
                });

                source.addEventListener('complete', function (e) {
                    clearInterval(timer);
                    source.close();
                    const result = JSON.parse(e.data);

                    // Mark all steps as done
                    if (progressBarFill) progressBarFill.style.width = '100%';
                    var stepEls = stepsContainer.querySelectorAll('.progress-step');
                    stepEls.forEach(function (el) {
                        el.classList.remove('active', 'pending');
                        el.classList.add('done');
                    });

                    // Short delay then show results
                    setTimeout(function () {
                        if (result.success) {
                            showResults(result);
                        } else {
                            showError(result.message);
                        }
                    }, 500);
                });

                source.addEventListener('error', function () {
                    clearInterval(timer);
                    source.close();
                    showError("La connexion au serveur a été perdue.");
                });
            })
            .catch(function (err) {
                clearInterval(timer);
                showError("Erreur de connexion au serveur: " + err.message);
            });
        });
    }

    // --- Update progress steps ---
    function updateProgress(stepName, pct, completedSteps, currentStep) {
        // Update progress bar
        if (progressBarFill) {
            progressBarFill.style.width = pct + '%';
        }

        // Mark previous step as done
        var existingSteps = stepsContainer.querySelectorAll('.progress-step');
        existingSteps.forEach(function (el) {
            if (el.classList.contains('active')) {
                el.classList.remove('active');
                el.classList.add('done');
            }
        });

        // Check if step already exists
        var stepId = 'step-' + pct;
        if (document.getElementById(stepId)) return;

        // Add new step
        var stepDiv = document.createElement('div');
        stepDiv.className = 'progress-step active';
        stepDiv.id = stepId;
        stepDiv.innerHTML = '<span class="step-icon"></span> ' + stepName;
        stepsContainer.appendChild(stepDiv);
    }

    // --- Show error ---
    function showError(message) {
        progressArea.classList.remove('active');
        if (uploadSection) uploadSection.style.display = 'none';

        resultsSection.innerHTML =
            '<div class="error-card">' +
                '<div class="error-title">Erreur</div>' +
                '<div class="error-message">' + escapeHtml(message) + '</div>' +
            '</div>' +
            '<div class="text-center">' +
                '<a href="/" class="btn btn-primary">Nouvelle Optimisation</a>' +
            '</div>';
        resultsSection.classList.add('active');
    }

    // --- Show results ---
    function showResults(result) {
        progressArea.classList.remove('active');
        if (uploadSection) uploadSection.style.display = 'none';

        var stats = result.stats;
        var html = '';

        // Verdict banner
        var prefRate = stats.pref_rate_float || 0;
        var vetoCount = stats.veto_count || 0;
        var verdictClass, verdictText;

        if (prefRate >= 95 && vetoCount === 0) {
            verdictClass = 'verdict-green';
            verdictText = 'Excellent résultat';
        } else if (prefRate >= 80 && vetoCount <= 2) {
            verdictClass = 'verdict-yellow';
            verdictText = 'Bon résultat';
        } else if (prefRate >= 60) {
            verdictClass = 'verdict-orange';
            verdictText = 'Résultat acceptable avec compromis';
        } else {
            verdictClass = 'verdict-red';
            verdictText = 'Résultat dégradé — vérifiez les données';
        }

        html += '<div class="verdict-banner ' + verdictClass + '">' + verdictText + '</div>';

        // Hero metric
        var heroClass;
        if (prefRate >= 95) heroClass = 'hero-green';
        else if (prefRate >= 80) heroClass = 'hero-yellow';
        else if (prefRate >= 60) heroClass = 'hero-orange';
        else heroClass = 'hero-red';

        html += '<div class="hero-metric ' + heroClass + '">' +
                    '<div class="hero-label">Taux de satisfaction global</div>' +
                    '<div class="hero-value">' + stats.pref_rate + '</div>' +
                    '<div class="hero-detail">' + stats.pref_count + ' / ' + stats.total_assignments + ' sessions préférées</div>' +
                '</div>';

        // KPI cards
        var vetoKpiClass = vetoCount === 0 ? 'kpi-success' : 'kpi-danger';
        html += '<div class="kpi-row">' +
                    '<div class="kpi-card">' +
                        '<div class="kpi-value">' + stats.fully_satisfied_count + '</div>' +
                        '<div class="kpi-label">100% satisfaits</div>' +
                        '<div class="kpi-sub">' + stats.fully_satisfied_pct + ' des élèves</div>' +
                    '</div>' +
                    '<div class="kpi-card">' +
                        '<div class="kpi-value">' + stats.mostly_satisfied_count + '</div>' +
                        '<div class="kpi-label">&ge; 3/5 satisfaits</div>' +
                        '<div class="kpi-sub">' + stats.mostly_satisfied_pct + ' des élèves</div>' +
                    '</div>' +
                    '<div class="kpi-card ' + vetoKpiClass + '">' +
                        '<div class="kpi-value">' + vetoCount + '</div>' +
                        '<div class="kpi-label">Vetos non respectés</div>' +
                        '<div class="kpi-sub">sur ' + stats.total_assignments + ' sessions</div>' +
                    '</div>' +
                '</div>';

        // Secondary info
        html += '<div class="secondary-info">' +
                    '<span>' + stats.students_processed + ' élèves</span>' +
                    '<span>' + stats.workshops_processed + ' ateliers</span>' +
                    '<span>' + (result.solve_time || '?') + 's</span>' +
                '</div>';

        // Download button
        html += '<div class="download-area">' +
                    '<a href="/download_result/' + encodeURIComponent(result.filename) + '" class="btn btn-success btn-lg">' +
                        'Télécharger le Planning (.xlsx)' +
                    '</a>' +
                '</div>';

        // Preferences chart
        html += '<div class="chart-card">' +
                    '<div class="chart-title">Distribution des préférences par élève</div>' +
                    '<div class="chart-subtitle">Nombre de sessions préférées obtenues sur ' + stats.max_prefs_possible + '</div>' +
                    '<div class="chart-container"><canvas id="prefsChart"></canvas></div>' +
                '</div>';

        // Category diversity chart (if data exists)
        var catDiv = stats.category_diversity_distribution;
        if (catDiv && Object.keys(catDiv).length > 0) {
            html += '<div class="chart-card">' +
                        '<div class="chart-title">Diversité catégorielle</div>' +
                        '<div class="chart-subtitle">Nombre de catégories différentes couvertes par chaque élève (poids: ' + stats.category_diversity_weight + ')</div>' +
                        '<div class="chart-container"><canvas id="categoryChart"></canvas></div>' +
                    '</div>';
        }

        // New optimization button
        html += '<div class="text-center mt-4">' +
                    '<a href="/" class="btn btn-primary">Nouvelle Optimisation</a>' +
                '</div>';

        resultsSection.innerHTML = html;
        resultsSection.classList.add('active');

        // Render charts
        renderPrefsChart(stats);
        if (catDiv && Object.keys(catDiv).length > 0) {
            renderCategoryChart(stats);
        }
    }

    // --- Chart colors (red -> green gradient) ---
    function getBarColor(index, total) {
        var colors = [
            'rgba(220, 38, 38, 0.7)',   // 0/5 red
            'rgba(234, 88, 12, 0.7)',    // 1/5 orange
            'rgba(217, 119, 6, 0.7)',    // 2/5 amber
            'rgba(101, 163, 13, 0.7)',   // 3/5 light green
            'rgba(22, 163, 74, 0.7)',    // 4/5 green
            'rgba(5, 150, 105, 0.7)'     // 5/5 deep green
        ];
        var borders = [
            'rgba(220, 38, 38, 1)',
            'rgba(234, 88, 12, 1)',
            'rgba(217, 119, 6, 1)',
            'rgba(101, 163, 13, 1)',
            'rgba(22, 163, 74, 1)',
            'rgba(5, 150, 105, 1)'
        ];
        var i = Math.min(index, colors.length - 1);
        return { bg: colors[i], border: borders[i] };
    }

    // --- Preferences Distribution Chart ---
    function renderPrefsChart(stats) {
        var canvas = document.getElementById('prefsChart');
        if (!canvas) return;

        var dist = stats.prefs_distribution;
        var total = stats.max_prefs_possible || 5;
        var labels = [];
        var data = [];
        var bgColors = [];
        var borderColors = [];

        for (var i = 0; i <= total; i++) {
            labels.push(i + '/' + total);
            data.push(dist[String(i)] || dist[i] || 0);
            var c = getBarColor(i, total);
            bgColors.push(c.bg);
            borderColors.push(c.border);
        }

        try {
            new Chart(canvas, {
                type: 'bar',
                data: {
                    labels: labels,
                    datasets: [{
                        label: "Nombre d'élèves",
                        data: data,
                        backgroundColor: bgColors,
                        borderColor: borderColors,
                        borderWidth: 1,
                        barPercentage: 0.7,
                        categoryPercentage: 0.8
                    }]
                },
                options: {
                    responsive: true,
                    maintainAspectRatio: false,
                    scales: {
                        y: {
                            beginAtZero: true,
                            title: { display: true, text: "Nombre d'élèves", font: { weight: 'bold' } },
                            ticks: { stepSize: 1 }
                        },
                        x: {
                            title: { display: true, text: 'Sessions préférées obtenues', font: { weight: 'bold' } }
                        }
                    },
                    plugins: {
                        legend: { display: false },
                        tooltip: { backgroundColor: 'rgba(0, 0, 0, 0.7)', titleFont: { weight: 'bold' } },
                        datalabels: {
                            anchor: 'end',
                            align: 'top',
                            font: { weight: 'bold', size: 11 },
                            formatter: function (value) { return value > 0 ? value : ''; }
                        }
                    }
                },
                plugins: [ChartDataLabels]
            });
        } catch (error) {
            // Fallback without datalabels plugin
            try {
                new Chart(canvas, {
                    type: 'bar',
                    data: {
                        labels: labels,
                        datasets: [{
                            label: "Nombre d'élèves",
                            data: data,
                            backgroundColor: bgColors,
                            borderColor: borderColors,
                            borderWidth: 1,
                            barPercentage: 0.7,
                            categoryPercentage: 0.8
                        }]
                    },
                    options: {
                        responsive: true,
                        maintainAspectRatio: false,
                        scales: {
                            y: {
                                beginAtZero: true,
                                title: { display: true, text: "Nombre d'élèves", font: { weight: 'bold' } },
                                ticks: { stepSize: 1 }
                            },
                            x: {
                                title: { display: true, text: 'Sessions préférées obtenues', font: { weight: 'bold' } }
                            }
                        },
                        plugins: {
                            legend: { display: false },
                            tooltip: { backgroundColor: 'rgba(0, 0, 0, 0.7)', titleFont: { weight: 'bold' } }
                        }
                    }
                });
            } catch (e2) {
                console.error("Erreur création graphique préférences:", e2);
            }
        }
    }

    // --- Category Diversity Chart ---
    function renderCategoryChart(stats) {
        var canvas = document.getElementById('categoryChart');
        if (!canvas) return;

        var catDivData = stats.category_diversity_distribution;
        var numCategories = (stats.categories || []).length;
        var labels = [];
        var data = [];

        for (var i = 0; i <= numCategories; i++) {
            var key = i + '/' + numCategories;
            labels.push(key);
            data.push(catDivData[key] || 0);
        }

        try {
            new Chart(canvas, {
                type: 'bar',
                data: {
                    labels: labels,
                    datasets: [{
                        label: "Nombre d'élèves",
                        data: data,
                        backgroundColor: 'rgba(5, 150, 105, 0.6)',
                        borderColor: 'rgba(5, 150, 105, 1)',
                        borderWidth: 1,
                        barPercentage: 0.7,
                        categoryPercentage: 0.8
                    }]
                },
                options: {
                    responsive: true,
                    maintainAspectRatio: false,
                    scales: {
                        y: {
                            beginAtZero: true,
                            title: { display: true, text: "Nombre d'élèves", font: { weight: 'bold' } },
                            ticks: { stepSize: 1 }
                        },
                        x: {
                            title: { display: true, text: 'Catégories couvertes', font: { weight: 'bold' } }
                        }
                    },
                    plugins: {
                        legend: { display: false },
                        tooltip: { backgroundColor: 'rgba(0, 0, 0, 0.7)', titleFont: { weight: 'bold' } }
                    }
                }
            });
        } catch (error) {
            console.error("Erreur création graphique catégories:", error);
        }
    }

    // --- Utility ---
    function escapeHtml(str) {
        var div = document.createElement('div');
        div.appendChild(document.createTextNode(str));
        return div.innerHTML;
    }
});

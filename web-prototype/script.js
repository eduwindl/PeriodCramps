/**
 * TechNet Reporting System - Frontend Simulation Logic
 * Simulates uploading files, parsing process, and document generation.
 */

document.addEventListener('DOMContentLoaded', () => {
    // Initialize Lucide Icons
    lucide.createIcons();

    // DOM Elements
    const fileVisitas = document.getElementById('file-visitas');
    const fileEquipos = document.getElementById('file-equipos');
    const btnGenerate = document.getElementById('btn-generate');
    
    const processStatus = document.getElementById('process-status');
    const progressBar = document.getElementById('progress-bar');
    const percentageText = document.getElementById('percentage');
    const statusText = document.getElementById('status-text');
    const processSteps = document.querySelectorAll('.process-steps li');
    
    const successView = document.getElementById('success-view');
    const successFileName = document.getElementById('success-file-name');
    const reportMonthInput = document.getElementById('report-month');
    const btnDownload = document.getElementById('btn-download');

    // State Variables
    let hasVisitas = false;
    let hasEquipos = false;

    // File Upload Handlers
    function handleFileUpload(inputElement, zoneId, isVisitas) {
        if (inputElement.files && inputElement.files.length > 0) {
            const fileName = inputElement.files[0].name;
            const zone = document.getElementById(zoneId);
            const label = zone.querySelector('.upload-content');
            const badge = zone.querySelector('.status-badge');
            
            // Validate extension
            if (!fileName.endsWith('.xlsx')) {
                alert('Por favor selecciona un archivo Excel (.xlsx)');
                inputElement.value = '';
                return;
            }

            // Update UI
            label.classList.add('uploaded');
            badge.classList.remove('pending');
            badge.classList.add('ready');
            badge.textContent = 'Cargado';

            // Change icon to show success
            const iconWrapper = zone.querySelector('.upload-icon-wrapper');
            iconWrapper.innerHTML = '<i data-lucide="check"></i>';
            lucide.createIcons();

            if (isVisitas) hasVisitas = true;
            else hasEquipos = true;

            // Enable generation button if both are uploaded
            if (hasVisitas && hasEquipos) {
                btnGenerate.removeAttribute('disabled');
            }
        }
    }

    fileVisitas.addEventListener('change', (e) => handleFileUpload(e.target, 'zone-visitas', true));
    fileEquipos.addEventListener('change', (e) => handleFileUpload(e.target, 'zone-equipos', false));

    // For demonstration: allow skipping upload by clicking the generate button if we force enable it
    // Uncomment lower line to test without uploading
    btnGenerate.removeAttribute('disabled');

    // Generate Report Workflow
    btnGenerate.addEventListener('click', async () => {
        // Disable everything
        btnGenerate.setAttribute('disabled', 'true');
        btnGenerate.style.display = 'none';
        
        // Setup Date info
        const dateVal = reportMonthInput.value; // e.g., "2026-02"
        const [year, month] = dateVal.split('-');
        
        // Show Process UI
        successView.style.display = 'none';
        processStatus.style.display = 'block';
        
        // Simulation timings
        const steps = [
            { text: "Cargando DataFrames de pandas...", duration: 1500, percent: 20, element: processSteps[0] },
            { text: "Limpiando datos y cruzando información...", duration: 2000, percent: 45, element: processSteps[1] },
            { text: "Calculando estadísticas (KPIs y agregaciones)...", duration: 1800, percent: 75, element: processSteps[2] },
            { text: "Renderizando tablas y gráficos en Word (docx)...", duration: 2500, percent: 100, element: processSteps[3] }
        ];

        let currentPercent = 0;

        // Async Sleep Helper
        const sleep = ms => new Promise(r => setTimeout(r, ms));

        // Step Runner
        for (let i = 0; i < steps.length; i++) {
            const step = steps[i];
            
            statusText.textContent = step.text;
            step.element.classList.remove('pending');
            step.element.classList.add('active');

            // Animate progress bar smoothly
            const percentDiff = step.percent - currentPercent;
            const stepDuration = step.duration;
            const interval = 50; // ms
            const ticks = stepDuration / interval;
            const percentPerTick = percentDiff / ticks;

            for(let t = 0; t < ticks; t++) {
                currentPercent += percentPerTick;
                progressBar.style.width = Math.min(Math.round(currentPercent), 100) + '%';
                percentageText.textContent = Math.min(Math.round(currentPercent), 100) + '%';
                await sleep(interval);
            }

            step.element.classList.remove('active');
            step.element.classList.add('done');
        }

        // Finish Success State
        statusText.textContent = "¡Operación completada con éxito!";
        statusText.style.color = "var(--success-color)";
        
        await sleep(500);
        
        processStatus.style.display = 'none';
        
        // Show Success Card
        successFileName.textContent = `reporte_${year}_${month}.docx`;
        successView.style.display = 'block';
    });

    // Reset/Download Mockup
    btnDownload.addEventListener('click', () => {
        // Trigger a fake download or just reset the UI
        alert("En el sistema real, esto descargará el archivo .docx localmente.");
        
        // Reset UI for next try
        setTimeout(() => {
            successView.style.display = 'none';
            btnGenerate.style.display = 'inline-flex';
            btnGenerate.removeAttribute('disabled');
            
            // resets
            progressBar.style.width = '0%';
            percentageText.textContent = '0%';
            statusText.textContent = 'Analizando datos...';
            statusText.style.color = 'var(--accent-color)';
            processSteps.forEach(el => {
                el.className = 'pending';
            });
        }, 1500);
    });
});

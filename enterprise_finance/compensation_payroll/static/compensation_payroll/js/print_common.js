
function printCommon(elementId) {
    const element = document.getElementById(elementId);

    if (!element) {
        alert('Content not found!');
        return;
    }

    const printWindow = window.open('', '_blank', 'width=800,height=600');

    if (!printWindow) {
        alert('Pop-up blocked. Please allow pop-ups for this site.');
        return;
    }

    const tableHtml = element.querySelector('table')?.outerHTML || '';
    const headingText = element.querySelector('h4')?.textContent || 'Summary';

    // Generate a unique document title (used by browser as default PDF name)
    const timestamp = new Date().toISOString().replace(/[:.]/g, '-'); // Safe for filenames
    const docTitle = `print_Summary_${timestamp}`;

    const htmlContent = `
        <!DOCTYPE html>
        <html>
        <head>
            <title>${docTitle}</title>
            <meta charset="UTF-8">
            <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/css/bootstrap.min.css" rel="stylesheet">
            <style>
                @page { size: auto; margin: 10mm; }
                body { padding: 15px; font-family: Arial, sans-serif; color: #000; }
                table { width: 100%; border-collapse: collapse; margin-bottom: 20px; }
                th, td { border: 1px solid #ddd; padding: 8px; }
                .bg-info { background-color: #d1e7ff !important; }
                .fw-bold { font-weight: bold !important; }
                .table-primary { background-color: #d1e7ff; }
                hr { border-top: 1px solid #000; }
                .print-header { margin-bottom: 20px; text-align: center; }
                .print-footer { margin-top: 20px; text-align: right; font-size: 0.9em; }
            </style>
        </head>
        <body>
            <div class="print-header">
                <h5>Print Summary</h5>
                <h6>${headingText}</h6>
            </div>
            ${tableHtml}
            <div class="print-footer">
                <p>Printed on: ${new Date().toLocaleDateString()}</p>
            </div>
            <script>
                window.onload = function() {
                    setTimeout(function() {
                        window.focus();
                        window.print();
                        setTimeout(function() {
                            window.close();
                        }, 300);
                    }, 200);
                };
            <\/script>
        </body>
        </html>
    `;

    printWindow.document.open();
    printWindow.document.write(htmlContent);
    printWindow.document.close();
}



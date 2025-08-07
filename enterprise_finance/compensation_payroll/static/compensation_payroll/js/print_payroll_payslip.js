

function printPayrollPayslip(elementId) {
    const element = document.getElementById(elementId);

    if (!element) {
        alert('Content not found!');
        return;
    }

    const printWindow = window.open('', '_blank', 'width=1000,height=800');

    if (!printWindow) {
        alert('Pop-up blocked. Please allow pop-ups for this site.');
        return;
    }

    const fullContentHtml = element.innerHTML;
    const headingText = element.querySelector('h4')?.textContent || 'Summary';
    const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
    const docTitle = `Compensation_And_Payroll_Summary_${timestamp}`;

    const htmlContent = `
        <!DOCTYPE html>
        <html>
        <head>
            <title>${docTitle}</title>
            <meta charset="UTF-8">
            <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/css/bootstrap.min.css" rel="stylesheet">
            <style>
                @page {
                    size: A4;
                    margin: 0;
                }

                html, body {
                    margin: 0;
                    padding: 10mm;
                    font-family: Arial, sans-serif;
                    font-size: 13px;
                    color: #000;
                    background: white;
                    overflow: hidden;
                    -webkit-print-color-adjust: exact;
                    print-color-adjust: exact;
                }

                *, *::before, *::after {
                    box-sizing: border-box;
                    max-width: 100%;
                    word-break: break-word;
                }

                .print-header {
                    text-align: center;
                    margin-bottom: 10px;
                }

                .print-body {
                    width: 100%;
                    page-break-inside: avoid;
                }

                .print-footer {
                    margin-top: 10px;
                    text-align: right;
                    font-size: 10px;
                    color: #555;
                }

                table {
                    width: 100%;
                    border-collapse: collapse;
                    table-layout: fixed;
                }

                th, td {
                    padding: 5px;
                    border: 1px solid #ccc;
                    text-align: left;
                }

                .row {
                    display: flex;
                    flex-wrap: wrap;
                }

                .col-md-6 {
                    width: 50%;
                    padding: 5px;
                }

                @media print {
                    html, body {
                        margin: 0 !important;
                        padding: 10mm !important;
                        overflow: hidden !important;
                    }

                    .no-print {
                        display: none !important;
                    }

                    .print-body, .row, .col-md-6, table {
                        page-break-inside: avoid !important;
                        break-inside: avoid !important;
                    }

                    a[href]::after {
                        content: none !important;
                    }
                }
            </style>
        </head>
        <body>
            <div class="print-header">
                <h3>${headingText}</h3>
            </div>
            <div class="print-body">
                ${fullContentHtml}
            </div>
            <div class="print-footer">
                Printed on: ${new Date().toLocaleString()}
            </div>

            <script>
                let printed = false;

                const beforePrint = () => { printed = true; };
                // Removed afterPrint event closing to avoid issues in Firefox

                window.addEventListener('beforeprint', beforePrint);
                // window.addEventListener('afterprint', () => { if (printed) window.close(); });

                window.matchMedia('print').addListener((mql) => {
                    if (mql.matches) beforePrint();
                    // else afterPrint(); // Removed for Firefox compatibility
                });

                window.addEventListener('load', function () {
                    setTimeout(() => {
                        window.focus();
                        window.print();

                        // Fallback: forcibly close window 7 seconds after print dialog opens
                        setTimeout(() => {
                            window.close();
                        }, 700);
                    }, 300);
                });
            <\/script>
        </body>
        </html>
    `;

    printWindow.document.open();
    printWindow.document.write(htmlContent);
    printWindow.document.close();
}


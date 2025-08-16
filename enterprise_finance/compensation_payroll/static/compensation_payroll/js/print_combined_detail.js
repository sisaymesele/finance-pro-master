

function printCombinedDetail(elementId) {
    const element = document.getElementById(elementId);

    if (!element) {
        alert('Content not found!');
        return;
    }

    const printDate = new Date();
    const timestamp = printDate.toISOString().replace(/[:.]/g, '-');
    const docTitle = `Compensation_And_Payroll_Detail_${timestamp}`;

    // Open a new blank window with unique URL to prevent replacing existing print window
    const printWindow = window.open('about:blank?' + timestamp, '_blank', 'width=800,height=1000');

    if (!printWindow) {
        alert('Pop-up blocked. Please allow pop-ups for this site.');
        return;
    }

    const fullContentHtml = element.innerHTML;

    const htmlContent = `
        <html>
        <head>
            <title>${docTitle}</title>
            <style>
                @media print {
                    html, body {
                        width: 210mm;
                        height: 297mm;
                        margin: 0;
                        padding: 0;
                        font-family: 'Arial', sans-serif;
                        color: #000;
                        background: #fff;
                        font-size: 12pt;
                    }

                    @page {
                        size: A4;
                        margin: 1.2cm;
                    }

                    .no-print {
                        display: none !important;
                    }

                    .print-wrapper {
                        width: 85% !important;
                        margin: 0 auto !important;
                        padding: 10px;
                    }

                    .card {
                        break-inside: avoid;
                        page-break-inside: avoid;
                        border: 1px solid #ccc;
                        padding: 10px;
                        margin-bottom: 12px;
                        box-shadow: none;
                        width: 100%;
                    }

                    .card-body {
                        padding: 0;
                    }

                    h2 {
                        font-size: 16pt;
                        margin-bottom: 10px;
                        text-align: center;
                    }

                    h5 {
                        font-size: 13pt;
                        margin-top: 15px;
                        margin-bottom: 10px;
                    }

                    h6 {
                        font-size: 11pt;
                        margin-top: 10px;
                        margin-bottom: 6px;
                    }

                    table {
                        width: 100%;
                        border-collapse: collapse;
                        margin-bottom: 10px;
                        break-inside: avoid;
                        page-break-inside: avoid;
                    }

                    th, td {
                        border: 1px solid #aaa;
                        padding: 4px 6px;
                        font-size: 9pt;
                    }

                    .table-responsive {
                        overflow: visible !important;
                    }

                    .row {
                        display: flex;
                        flex-wrap: nowrap;
                        gap: 16px;
                    }

                    .border-end {
                        border-right: 1px solid #ccc !important;
                        padding-right: 10px !important;
                    }

                    .print-timestamp {
                        text-align: right;
                        font-size: 9pt;
                        margin-bottom: 8px;
                        color: #555;
                    }
                }

                body {
                    font-family: 'Arial', sans-serif;
                }
            </style>
        </head>
        <body>
            <div class="print-wrapper">
                <div class="print-timestamp">Printed on: ${printDate.toLocaleString()}</div>
                ${fullContentHtml}
            </div>

            <script>
                window.onload = function() {
                    window.focus();
                    window.print();
                    setTimeout(() => window.close(), 700);
                };
            </script>
        </body>
        </html>
    `;

    printWindow.document.open();
    printWindow.document.write(htmlContent);
    printWindow.document.close();
}





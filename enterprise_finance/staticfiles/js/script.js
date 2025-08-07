

function printContent(contentId, title, customStyles = "") {
    // Get content by ID or fallback to a default container
    const content = document.getElementById(contentId)?.innerHTML || document.querySelector('.container')?.innerHTML;

    // Return if no content is found
    if (!content) return;

    // Open a new window to print the content
    const printWindow = window.open("", "", "width=800,height=600");

    // Write the HTML content, including title, Bootstrap, custom styles, and the main content
    printWindow.document.write(`
        <html>
        <head>
            <title>${title}</title>
            <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css">
            <style>
                body {
                    font-size: 12px;
                    margin: 0;
                    padding: 0;
                    height: 100%;
                    overflow: hidden;
                }
                .table {
                    border-collapse: collapse;
                    width: 100%;
                    margin: 5px auto;
                }
                .table th, .table td {
                    border: 1px solid #000;
                    text-align: left;
                    padding: 5px;
                    font-size: 14px;
                }
                .d-flex { display: flex; flex-wrap: wrap; justify-content: space-between; }
                .border { border: 1px solid #000 !important; }
                .card-body {
                    page-break-inside: avoid;
                    padding: 10px;
                    margin: 0;
                }

                /* Avoid page breaks */
                @page {
                    size: A4;
                    margin: 10mm;
                }

                @media print {
                    body {
                        font-size: 10px;
                        margin: 0;
                        padding: 0;
                        height: 100%;
                    }
                    .container, .card, .card-body {
                        max-width: 100%;
                        padding: 5px;
                        margin: 0;
                        page-break-inside: avoid;
                        font-size: 14pt; /* Increase base font size for printing */
                    }
                    .d-flex {
                        flex-wrap: nowrap;
                    }
                    .table {
                        width: 100%;
                        page-break-before: auto;
                        page-break-after: auto;
                        font-size: 14pt; /* Increase base font size for printing */
                    }
                }

                ${customStyles}
            </style>
        </head>
        <body class="container p-4">
            <div class="card bg-light p-4 text-center">
                <div class="card-header">
                    <h4>${title}</h4>
                </div>
                <div class="card-body">
                    <div class="container">
                        ${content}
                    </div>
                </div>
            </div>
        </body>
        </html>
    `);
    printWindow.document.close();

    // Delay the print action to ensure content is loaded
    setTimeout(() => printWindow.print(), 100);
}


// Usage examples:
function printPayrollPayslip() { printContent(); }
function printPayrollMonthSummary(yearMonth) { printContent("regular-payroll-summary-" + yearMonth, "Regular Payroll Summary"); }
function printPayrollYearSummary(year) { printContent("yearly-payroll-summary-" + year, "Yearly Payroll Summary"); }
//function printRegularPayrollJournalEntry() { printContent("", "Payroll Journal Entry"); }

function printMultiMonthBonusDetail() { printContent("", "Multi-Month Bonus Detail"); }
function printAnnualLeavePayDetail() { printContent(); }
function printSeverancePayDetail() { printContent(); }

//monthly journal entry
function printRegularPayrollJournalEntry(elementId, title) {
    const printContent = document.getElementById(elementId).innerHTML;
    const printWindow = window.open('', '_blank', 'width=800,height=600');

    printWindow.document.write(`
        <html>
            <head>
                <title>${title}</title>
                <style>
                    body {
                        font-family: Arial, sans-serif;
                        margin: 20px;
                    }
                    h4 {
                        text-align: center;
                        margin-bottom: 20px;
                    }
                    table {
                        width: 100%;
                        border-collapse: collapse;
                        margin-bottom: 20px;
                        font-size: 14pt; /* Increase base font size for printing */
                    }
                    th, td {
                        border: 1px solid #000;
                        padding: 12px; /* Increased padding for better spacing */
                        text-align: left;
                    }
                    th {
                        background-color: #f2f2f2;
                        font-weight: bold;
                    }
                    .print-header {
                        text-align: center;
                        margin-bottom: 20px;
                    }
                    .print-footer {
                        text-align: center;
                        margin-top: 20px;
                        font-size: 12pt; /* Larger footer text size */
                        color: #555;
                    }
                </style>
            </head>
            <body>
                <div class="print-header">
                    <h2>Payroll Journal Entry</h2>
                    <p>${title}</p>
                </div>
                ${printContent}
                <div class="print-footer">
                    <p>Generated on: ${new Date().toLocaleDateString()}</p>
                </div>
                <script>
                    window.onload = function() {
                        window.print();
                        window.close();
                    };
                <\/script>
            </body>
        </html>
    `);
    printWindow.document.close();
}
//multimonth summary

function printMultiMonthBonusSummary(elementId, title) {
   const printContent = document.getElementById(elementId).innerHTML;
   const printWindow = window.open('', '_blank', 'width=800,height=600');

   printWindow.document.write(`
       <html>
           <head>
               <title>${title}</title>
               <style>
                   body {
                        font-family: Arial, sans-serif;
                        display: flex;
                        justify-content: center;
                        align-items: center;
                        height: 30vh;
                        margin: 0;
                        flex-direction: column;
                    }
                    .printable-area {
                        width: 80%;
                        max-width: 800px;
                        padding: 20px;
                        border: 1px solid #ccc;
                        background-color: #fff;
                        box-shadow: 0 0 10px rgba(0, 0, 0, 0.1);
                    }
                    table {
                        width: 100%;
                        border-collapse: collapse;
                        margin: 20px 0;
                        font-size: 14pt; /* Increase base font size for printing */
                    }
                    th, td {
                        border: 1px solid #ddd;
                        padding: 8px;
                        text-align: left;
                    }
                    th {
                        background-color: #f2f2f2;
                    }
                    .text-center {
                        text-align: center;
                    }
                    .fw-bold {
                        font-weight: bold;
                    }
                    @media print {
                        /* Ensure the button is not displayed in print mode */
                        button {
                            display: none;
                        }
                    }
               </style>
           </head>
           <body>
               <div>
                   <h2 class="text-center">${title}</h2>
                   ${printContent}
               </div>
               <script>
                   window.onload = function() {
                       window.print();
                       window.close();
                   };
               <\/script>
           </body>
       </html>
   `);
   printWindow.document.close();
}
//multimonth summary
//multimonth journal entry

function printMultiMonthBonusJournalEntrySummary(elementId, month, year) {
    const printContent = document.getElementById(elementId).innerHTML;
    const printWindow = window.open('', '_blank', 'width=800,height=600');

    printWindow.document.write(`
        <html>
            <head>
                <title>Journal Entry Summary for ${month} - ${year}</title>
                <style>
                    body {
                        font-family: Arial, sans-serif;
                        display: flex;
                        justify-content: center;
                        align-items: center;
                        height: 30vh;
                        margin: 0;
                        flex-direction: column;
                        font-size: 14pt; /* Increase base font size for printing */
                    }
                    .printable-area {
                        width: 80%;
                        max-width: 800px;
                        padding: 20px;
                        border: 1px solid #ccc;
                        background-color: #fff;
                        box-shadow: 0 0 10px rgba(0, 0, 0, 0.1);
                    }
                    table {
                        width: 100%;
                        border-collapse: collapse;
                        margin: 20px 0;
                        font-size: 14pt; /* Increase base font size for printing */
                    }
                    th, td {
                        border: 1px solid #ddd;
                        padding: 8px;
                        text-align: left;
                    }
                    th {
                        background-color: #f2f2f2;
                    }
                    .text-center {
                        text-align: center;
                    }
                    .fw-bold {
                        font-weight: bold;
                    }
                    @media print {
                        /* Ensure the button is not displayed in print mode */
                        button {
                            display: none;
                        }
                    }
                </style>
            </head>
            <body>
                <div class="print-container">
                    ${printContent}
                </div>
                <script>
                    window.onload = function() {
                        window.print();
                        window.close();
                    };
                <\/script>
            </body>
        </html>
    `);
    printWindow.document.close();
}

//multimonth journal entry

//annual leave pay summary

function printAnnualLeavePaySummary(elementId) {
    // Get the content to print
    const printContents = document.getElementById(elementId).innerHTML;

    // Open a new pop-up window
    const printWindow = window.open('', '_blank', 'width=800,height=600');

    // Write the content to the pop-up window
    printWindow.document.write(`
        <html>
            <head>
                <title>Print</title>
                <style>
                    body {
                        font-family: Arial, sans-serif;
                        display: flex;
                        justify-content: center;
                        align-items: center;
                        height: 30vh;
                        margin: 0;
                        flex-direction: column;
                    }
                    .printable-area {
                        width: 80%;
                        max-width: 800px;
                        padding: 20px;
                        border: 1px solid #ccc;
                        background-color: #fff;
                        box-shadow: 0 0 10px rgba(0, 0, 0, 0.1);
                    }
                    table {
                        width: 100%;
                        border-collapse: collapse;
                        margin: 20px 0;
                        font-size: 14pt; /* Increase base font size for printing */

                    }
                    th, td {
                        border: 1px solid #ddd;
                        padding: 8px;
                        text-align: left;
                    }
                    th {
                        background-color: #f2f2f2;
                    }
                    .text-center {
                        text-align: center;
                    }
                    .fw-bold {
                        font-weight: bold;
                    }
                    @media print {
                        /* Ensure the button is not displayed in print mode */
                        button {
                            display: none;
                        }
                    }
                </style>
            </head>
            <body>
                <div class="printable-area">
                    ${printContents}
                </div>
                <script>
                    // Automatically print and close the window after loading
                    window.onload = function() {
                        window.print();
                        window.close();
                    };
                <\/script>
            </body>
        </html>
    `);

    // Close the document for writing
    printWindow.document.close();
}
//annual leave pay summary
//annual leve pay journal

function printAnnualLeavePayJournalEntry(elementId, month, year) {
    // Get the content to print (no button removal)
    const printContent = document.getElementById(elementId).cloneNode(true);

    // Open a new pop-up window
    const printWindow = window.open('', '_blank', 'width=800,height=600');

    // Write the content to the pop-up window
    printWindow.document.write(`
        <html>
            <head>
                <title>Journal Entry Summary</title>
                <style>
                    body {
                        font-family: Arial, sans-serif;
                        display: flex;
                        justify-content: center;
                        align-items: center;
                        height: 30vh;
                        margin: 0;
                        flex-direction: column;
                    }
                    .printable-area {
                        width: 80%;
                        max-width: 800px;
                        padding: 20px;
                        border: 1px solid #ccc;
                        background-color: #fff;
                        box-shadow: 0 0 10px rgba(0, 0, 0, 0.1);
                    }
                    table {
                        width: 100%;
                        border-collapse: collapse;
                        margin: 20px 0;
                        font-size: 14pt; /* Increase base font size for printing */
                    }
                    th, td {
                        border: 1px solid #ddd;
                        padding: 8px;
                        text-align: left;
                    }
                    th {
                        background-color: #f2f2f2;
                    }
                    .text-center {
                        text-align: center;
                    }
                    .fw-bold {
                        font-weight: bold;
                    }
                    @media print {
                        /* Ensure the button is not displayed in print mode */
                        button {
                            display: none;
                        }
                    }
                </style>
            </head>
            <body>

                ${printContent.innerHTML}
                <script>
                    // Automatically print and close the window after loading
                    window.onload = function() {
                        window.print();
                        window.close();
                    };
                </script>
            </body>
        </html>
    `);

    // Close the document for writing
    printWindow.document.close();
}


//annual leve pay journal
//// severance summary
function printSeverancePaySummary(elementId, month, year) {
    // Get the content to print (no button removal)
    const printContent = document.getElementById(elementId).cloneNode(true);

    // Open a new pop-up window
    const printWindow = window.open('', '_blank', 'width=800,height=600');

    // Write the content to the pop-up window
    printWindow.document.write(`
        <html>
            <head>
                <title>Severance Pay Summary for ${month} ${year}</title>
                <style>
                    body {
                        font-family: Arial, sans-serif;
                        display: flex;
                        justify-content: center;
                        align-items: center;
                        height: 30vh;
                        margin: 0;
                        flex-direction: column;
                    }
                    .printable-area {
                        width: 80%;
                        max-width: 800px;
                        padding: 20px;
                        border: 1px solid #ccc;
                        background-color: #fff;
                        box-shadow: 0 0 10px rgba(0, 0, 0, 0.1);
                    }
                    table {
                        width: 100%;
                        border-collapse: collapse;
                        margin: 20px 0;
                        font-size: 14pt; /* Increase base font size for printing */
                    }
                    th, td {
                        border: 1px solid #ddd;
                        padding: 8px;
                        text-align: left;
                    }
                    th {
                        background-color: #f2f2f2;
                    }
                    .text-center {
                        text-align: center;
                    }
                    .fw-bold {
                        font-weight: bold;
                    }
                    @media print {
                        /* Ensure the button is not displayed in print mode */
                        button {
                            display: none;
                        }
                    }
               </style>
            </head>
            <body>

                ${printContent.innerHTML}
                <script>
                    // Automatically print and close the window after loading
                    window.onload = function() {
                        window.print();
                        window.close();
                    };
                <\/script>
            </body>
        </html>
    `);

    // Close the document for writing
    printWindow.document.close();
}


//severance pay summary

//severance pay journal entry
//

function printSeverancePayJournalEntry(elementId, month, year) {
    // Get the content to print
    const printContent = document.getElementById(elementId).cloneNode(true);

    // Open a new pop-up window
    const printWindow = window.open('', '_blank', 'width=800,height=600');

    // Write the content to the pop-up window
    printWindow.document.write(`
        <html>
            <head>
                <title>Severance Pay Journal Entry for ${month} ${year}</title>
                <style>
                    /* General styles */
                    body {
                        font-family: Arial, sans-serif;
                        margin: 0;
                        padding: 20px;
                        background-color: #f8f9fa;
                    }
                    .printable-area {
                        width: 100%;
                        max-width: 800px;
                        margin: 0 auto;
                        padding: 20px;
                        border: 1px solid #ddd;
                        background-color: #fff;
                        box-shadow: 0 0 10px rgba(0, 0, 0, 0.1);
                        border-radius: 8px;
                    }
                    .card {
                        border: 1px solid #ddd;
                        border-radius: 8px;
                        margin-bottom: 20px;
                        box-shadow: 0 0 10px rgba(0, 0, 0, 0.1);
                    }
                    .card-header {
                        background-color: #f8f9fa;
                        padding: 15px;
                        border-bottom: 1px solid #ddd;
                        border-top-left-radius: 8px;
                        border-top-right-radius: 8px;
                    }
                    .card-body {
                        padding: 20px;
                    }
                    table {
                        width: 100%;
                        border-collapse: collapse;
                        margin: 20px 0;
                        font-size: 14pt; /* Increase base font size for printing */
                    }
                    th, td {
                        border: 1px solid #ddd;
                        padding: 8px;
                        text-align: left;
                    }
                    th {
                        background-color: #f2f2f2;
                    }
                    .text-center {
                        text-align: center;
                    }
                    .fw-bold {
                        font-weight: bold;
                    }

                    @media print {
                        /* Ensure the button is not displayed in print mode */
                        button {
                            display: none;
                        }
                        /* Adjust printable area for print */
                        .printable-area {
                            box-shadow: none;
                            border: none;
                            padding: 0;
                        }
                        .card {
                            box-shadow: none;
                            border: none;
                        }
                    }
                </style>
            </head>
            <body>
                <div class="printable-area">
                    <div class="card">
                        <div class="card-header">
                            <h3 class="text-center fw-bold">Severance Pay Journal Entry</h3>
                        </div>
                        <div class="card-body">
                            ${printContent.innerHTML}
                        </div>
                    </div>
                </div>
                <script>
                    // Automatically print and close the window after loading
                    window.onload = function() {
                        window.print();
                        window.close();
                    };
                <\/script>
            </body>
        </html>
    `);

    // Close the document for writing
    printWindow.document.close();
}

//severance pay journal entry

//monthly compensation
//
function printMonthlyCompensationSummary(month, year) {
    const elementId = `monthly-compensation-summary-${month}-${year}`;
    const originalContent = document.getElementById(elementId);

    if (!originalContent) {
        alert("Content not found!");
        return;
    }

    // Open a new print window
    const printWindow = window.open('', '_blank', 'width=800,height=600');

    const htmlContent = `
        <html>
        <head>
            <title>Compensation Summary for ${month} ${year}</title>
            <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css">
            <style>
                body { font-family: Arial, sans-serif; padding: 20px; }
                .container { max-width: 800px; margin: auto; text-align: center; }
                .card { margin-top: 20px; padding: 20px; border: 1px solid #ddd; }
                @media print {
                        /* Ensure the button is not displayed in print mode */
                        button {
                            display: none;
                        }
                    }
            </style>
        </head>
        <body>
            <div class="container">
                <div class="card">
                    <h2 class="text-center">Compensation Summary</h2>
                    <h4 class="text-center">${month} ${year}</h4>
                    <hr>
                    ${originalContent.innerHTML}
                </div>
            </div>
        </body>
        </html>
    `;

    printWindow.document.write(htmlContent);
    printWindow.document.close();

    // Wait for content to load, then print
    printWindow.onload = function() {
        printWindow.print();
        printWindow.close();
    };
}



//monthly compensation
//yeary compensation

function printYearlyCompensationSummary(year) {
    const elementId = `yearly-compensation-summary-${year}`;
    const originalContent = document.getElementById(elementId);

    if (!originalContent) {
        alert("Content not found!");
        return;
    }

    // Open a new print window
    const printWindow = window.open('', '_blank', 'width=800,height=600');

    const htmlContent = `
        <html>
        <head>
            <title>Compensation Summary for ${year}</title>
            <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css">
            <style>
                body { font-family: Arial, sans-serif; padding: 20px; }
                .container { max-width: 800px; margin: auto; text-align: center; }
                .card { margin-top: 20px; padding: 20px; border: 1px solid #ddd; }
                @media print {
                        /* Ensure the button is not displayed in print mode */
                        button {
                            display: none;
                        }
                    }
            </style>
        </head>
        <body>
            <div class="container">
                <div class="card">
                    <h2 class="text-center">Compensation Summary</h2>
                    <h4 class="text-center">${year}</h4>
                    <hr>
                    ${originalContent.innerHTML}
                </div>
            </div>
        </body>
        </html>
    `;

    printWindow.document.write(htmlContent);
    printWindow.document.close();

    // Wait for content to load, then print
    printWindow.onload = function() {
        printWindow.print();
        printWindow.close();
    };
}

//yearly compensation


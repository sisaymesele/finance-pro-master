
//print earning monthly

function printAdjustment(divId) {
    var printContents = document.getElementById(divId).outerHTML;
    var newWindow = window.open('', '', 'width=900,height=700');
    newWindow.document.write(`
        <html>
            <head>
                <title>Monthly Earnings Adjustment</title>
                <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
                <style>
                    body { font-family: Arial, sans-serif; padding: 20px; }
                    .card-header { font-size: 1.5rem; }
                    table { font-size: 0.95rem; }
                </style>
            </head>
            <body>
                ${printContents}
            </body>
        </html>
    `);
    newWindow.document.close();
    newWindow.print();
}



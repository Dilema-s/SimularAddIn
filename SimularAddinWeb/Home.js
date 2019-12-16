(function() {
    "use strict";

    var cellToHighlight;
    var messageBanner;

    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function(reason) {
        $(document).ready(function () {

            // Initialize the FabricUI notification mechanism and hide it
            var element = document.querySelector('.ms-MessageBanner');
            messageBanner = new fabric.MessageBanner(element);
            messageBanner.hideBanner();

            // If not using Excel 2016, use fallback logic.
            if (!Office.context.requirements.isSetSupported('ExcelApi', '1.1')) {
                $('#button-text').text("Display!");
                $('#button-desc').text("Display the selection");

                $('#highlight-button').click(displaySelectedCells);
                return;
            }

            $('#button-text').text("QuantLib");
            $('#button-desc').text("Highlights the largest number.");

            //loadSampleData();

            // Add a click event handler for the highlight button.
            $('#highlight-button').click(quantlibTest);
            $('#generar-datos').click(loadSampleData);
            $('#ordenar-datos').click(sorterTable);
            $('#calculate-mean').click(calculateMean);
            $('#calculate-nd').click(NormalDistribution);
        });
    };

    function loadSampleData() {
        // Run a batch operation against the Excel object model
        Excel.run(function(ctx) {

                var sheet = ctx.workbook.worksheets.getActiveWorksheet();
                var expensesTable = sheet.tables.add("A1:B1", true /*hasHeaders*/);
                expensesTable.name = "ExpensesTable";

                expensesTable.getHeaderRowRange().values = [["Employee Name", "Employee Ratings"]];

                for (var i = 0; i <= 99; i++) {
                    expensesTable.rows.add(null /*add rows to the end of the table*/, [
                        ["Employee" + " " + i, Math.floor(Math.random() * (50 - 10 + 1)) + 10]
                    ]);
                }

                if (Office.context.requirements.isSetSupported("ExcelApi", "1.2")) {
                    sheet.getUsedRange().format.autofitColumns();
                    sheet.getUsedRange().format.autofitRows();
                }

                sheet.activate();
                return ctx.sync();
            })
            .catch(errorHandler);
    }

    function sorterTable() {
        Excel.run(function (ctx) {
            var sheet = ctx.workbook.worksheets.getActiveWorksheet();
            var expensesTable = sheet.tables.getItem("ExpensesTable");

            // Queue a command to sort data by the fourth column of the table (descending)
            var sortRange = expensesTable.getDataBodyRange();
            sortRange.sort.apply([
                {
                    key: 1,
                    ascending: true
                }
            ]);

            // Sync to run the queued command in Excel
            return ctx.sync();
        }).catch(errorHandler);
    }

    function calculateMean() {
        Excel.run(function (ctx) {
            var sheet = ctx.workbook.worksheets.getActiveWorksheet();

            var data = [
                ["Mean","=AVERAGE(B2:B101)"],
                ["Standard Deviation","=STDEV(B2:B101)"]
            ];

            var range = sheet.getRange("E10:F11");
            range.formulas = data;
            range.format.autofitColumns();

            // Sync to run the queued command in Excel
            return ctx.sync();
        }).catch(errorHandler);
      
    }

    function NormalDistribution() {
        Excel.run(function (ctx) {
            var sheet = ctx.workbook.worksheets.getActiveWorksheet();

            var data = [];
            var mean = sheet.getRange("F10").values;
            var sd = sheet.getRange("F11").values;

            for (var i = 2; i <= 101; i++) {
                data.push(["=NORM.DIST(B" + i + "," +mean +","+ sd +",FALSE)"]);
            }


            var range = sheet.getRange("C2:C101");
            range.formulas = data;
            range.format.autofitColumns();

            // Sync to run the queued command in Excel
            return ctx.sync();
        }).catch(errorHandler);
    }

    function createChart() {
        Excel.run(function (context) {
            var sheet = context.workbook.worksheets.getItem("Sample");
            var dataRange = sheet.getRange("A1:B13");
            var chart = sheet.charts.add("Line", dataRange, "auto");

            chart.title.text = "Sales Data";
            chart.legend.position = "right";
            chart.legend.format.fill.setSolidColor("white");
            chart.dataLabels.format.font.size = 15;
            chart.dataLabels.format.font.color = "black";

            return context.sync();
        }).catch(errorHandler);
    }

    function hightlightHighestValue() {
        // Run a batch operation against the Excel object model
        Excel.run(function(ctx) {
                // Create a proxy object for the selected range and load its properties
                var sourceRange = ctx.workbook.getSelectedRange().load("values, rowCount, columnCount");

                // Run the queued-up command, and return a promise to indicate task completion
                return ctx.sync()
                    .then(function() {
                        var highestRow = 0;
                        var highestCol = 0;
                        var highestValue = sourceRange.values[0][0];

                        // Find the cell to highlight
                        for (var i = 0; i < sourceRange.rowCount; i++) {
                            for (var j = 0; j < sourceRange.columnCount; j++) {
                                if (!isNaN(sourceRange.values[i][j]) && sourceRange.values[i][j] > highestValue) {
                                    highestRow = i;
                                    highestCol = j;
                                    highestValue = sourceRange.values[i][j];
                                }
                            }
                        }

                        cellToHighlight = sourceRange.getCell(highestRow, highestCol);
                        sourceRange.worksheet.getUsedRange().format.fill.clear();
                        sourceRange.worksheet.getUsedRange().format.font.bold = false;

                        // Highlight the cell
                        cellToHighlight.format.fill.color = "orange";
                        cellToHighlight.format.font.bold = true;
                    })
                    .then(ctx.sync);
            })
            .catch(errorHandler);
    }

    function quantlibTest() {
        // Run a batch operation against the Excel object model
        Excel.run(function (ctx) {
                function CallHandlerSync(inputUrl, inputdata) {
                    $.ajax({
                        url: inputUrl,
                        contentType: "application/json; charset=utf-8",
                        data: inputdata,
                        async: false,
                        success: function (result) {
                            var objResult = $.parseJSON(result);
                            showNotification(objResult.Total);
                        },
                        error: function (e, ts, et) {
                            showNotification("Error...");
                        }
                    });
                }
                return ctx.sync()
                    .then(function () {

                        

                        var number1 = 0.2;
                        CallHandlerSync("Handler.ashx?RequestType=Derivative", { 'number1': number1});
                    })
                    .then(ctx.sync);
            })
            .catch(errorHandler);
    }


    function displaySelectedCells() {
        Office.context.document.getSelectedDataAsync(Office.CoercionType.Text,
            function(result) {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    showNotification('The selected text is:', '"' + result.value + '"');
                } else {
                    showNotification('Error', result.error.message);
                }
            });
    }

    // Helper function for treating errors
    function errorHandler(error) {
        // Always be sure to catch any accumulated errors that bubble up from the Excel.run execution
        showNotification("Error", error);
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    }

    // Helper function for displaying notifications
    function showNotification(header, content) {
        $("#notification-header").text(header);
        $("#notification-body").text(content);
        messageBanner.showBanner();
        messageBanner.toggleExpansion();
    }
})();


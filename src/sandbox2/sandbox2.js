(function () {
    "use strict";

    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {
            
            // If not using Excel 2016, use fallback logic.
            // if (!Office.context.requirements.isSetSupported('ExcelApi', '1.1')) {
            //     $("#template-description").text("This sample will display the value of the cells that you have selected in the spreadsheet.");
            //     $('#button-text').text("Display!");
            //     $('#button-desc').text("Display the selection");

            //     $('#random-interval-button').click(displaySelectedCells);
            //     return;
            // }

            $("#template-description").text("This sample run random multiple times.");
            $('#button-text').text("Random interval!");
            $('#button-desc').text("Jus put values every second.");

            // Add a click event handler for the highlight button.
            $('#random-interval-button').click(intervalOnFirstCell);
        });
    };



    function intervalOnFirstCell() {
        setInterval(() => {
          var values = [
              [Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000)],
              [Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000)],
              [Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000)]
          ];
      
          // Run a batch operation against the Excel object model
          Excel.run(function (ctx) {
              var sheet = ctx.workbook.worksheets.getActiveWorksheet();
              sheet.getRange("B3:D5").values = values;
              sheet.getRange("B2:D2").values = [['A', 'B', 'C']];
              return ctx.sync();
          })
          .catch(errorHandler);
        }, 100);
      };

    // Helper function for treating errors
    function errorHandler(error) {
        // Always be sure to catch any accumulated errors that bubble up from the Excel.run execution
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    };

})();

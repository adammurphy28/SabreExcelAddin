(function ($) {
    "use strict";

    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function () {

        $(document).ready(async function () {

            const copyArea = $('.copyArea'),
                info = await getCrewInfo();
            let m = 0;

            console.log(info);

            // Loop through filtered items
            $.each(info, function (key, value) {

                const memberName = value[0],
                    memberHtml = $(`<div class="selection"><input id="member-${m}" class="member" name="member-${m}" type="checkbox" /><label for="member-${m}">${memberName}</label></div>`);

                $('.crew-member-select-container').append(memberHtml);

                m++;
            });

            // Loop through input containers
            $('.selection').each(function () {

                // Find inputs for members
                const input = $(this).find('input.member'),
                    id = input.attr('id');

                let position = '';

                // For some reason not all member id's are strings???
                if (typeof (id) === 'string') {

                    position = parseInt(id.replace('member-', ''));
                }

                // When clicking member
                input.on('click', function () {

                    const lineBreak = '\n';

                    // Start with Name
                    let sabreSymbol = '§',
                        sabreFormatting = Object.values(info[position])[0] + sabreSymbol;

                    // Loop through each value for that row
                    $.each(info[position], function (key, value) {

                        // If value contains the info we need
                        if (value.includes("3DOCS/DB") || value.includes("3CTCE") || value.includes("3CTCM") || value.includes("3DOCO") || value.includes("3DOCS/P/")) {

                            // Append value to formatted string
                            sabreFormatting += value + sabreSymbol;
                        }
                    });

                    // If copy area is empty
                    if (copyArea.text() === '') {

                        // Append string as is
                        copyArea.append(sabreFormatting);
                    } else {

                        // Else append string with linebreak added
                        copyArea.append(lineBreak + sabreFormatting);
                    }
                })
            });

            // Select all inputs
            $('#selectAll').on('click', selectAll);

            // Copy text from textarea on click
            copyArea.on('click', function () {
                navigator.clipboard.writeText($(this).text());
            })
        });
    };

    async function getCrewInfo() {

        let filteredItems = {};

        await Excel.run(async (context) => {

            // Get Text from cells that contain values
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            const usedCells = sheet.getUsedRange(true);

            usedCells.load("text");

            await context.sync();

            const cellsWithText = usedCells.text;
            let i = 0;

            // Loop through each row
            $.each(cellsWithText, function (key, value) {

                let values = [];

                // Loop through each cell in the row
                $.each(value, function (key, nextValue) {

                    // If cell is not empty
                    if (nextValue !== "" && nextValue !== null && nextValue !== '') {

                        // Transform everything to uppercase
                        let valueFormatted = nextValue.toUpperCase();

                        // Remove spaces at beginning of cell
                        if (valueFormatted.charAt(0) === " ") {
                            valueFormatted = valueFormatted.substring(1, valueFormatted.length);
                        }

                        // Remove spaces at end of cell
                        if (valueFormatted.charAt(valueFormatted.length) === " ") {
                            valueFormatted = valueFormatted.substring(0, valueFormatted.length - 1);
                        }

                        // Pushed formatted string to array
                        values.push(valueFormatted);
                    }
                });

                // Remove any section headers
                if (values.length > 1 && !values.includes("NAME")) {
                    filteredItems[i] = values;
                    i++;
                }
            });
        }).catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });

        return filteredItems;
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

    function selectAll() {

        $('.selection').each(function () {

            const input = $(this).find('input');

            if ($('#selectAll').prop('checked') === true) {
                input.prop('checked', true);
            } else {
                input.prop('checked', false);
            }
        });
    }

}(this.jQuery));

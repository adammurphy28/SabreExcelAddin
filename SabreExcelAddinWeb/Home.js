(function ($) {
    "use strict";

    // TODO
    /*
        - Add Refresh Button
        - Make Select All add everything to copyarea
        - Add Frequent Flyer Dropdown
        - Remove items from textarea on deselect
        - Add Styling
        - Add Toggle for Passport
        - Add Toggle for Meal Option:
        (AVMLA,BBMLA,BLMLA,CHMLA,DBMLA,FPMLA,GFMLA,HNMLA,KSMLA,LCMLA,LFMLA,LSMLA,MOMLA,NLMLA,NOMLA,RVMLA,SFMLA,SPMLA,VGMLA,VJMLA,VLMLA,VOMLA)
        - Add iteration toggle(Start with -2.1, -3.1, etc.)
    */

    // To account for more than 1 member per record
    let crewIteration = 1;

    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function () {

        $(document).ready(async function () {

            const copyArea = $('.copyArea'),
                crewInfo = await getCrewInfo();
            let m = 0;

            console.log(crewInfo);

            // Loop through filtered items
            $.each(crewInfo, function (key, value) {

                const memberName = value[0],
                    memberHtml = $(`<div class="selection"><input id="member-${m}" class="member" name="member-${m}" type="checkbox" /><label for="member-${m}">${memberName}</label></div>`);

                $('.crew-member-select-container').append(memberHtml);

                m++;
            });

            // When clicking member
            $('input.member').on('click', function () {
                formatInfo($(this), crewInfo, copyArea);
                crewIteration++;
            });

            // Select all inputs
            $('#selectAll').on('click', selectAll);

            // Copy text from textarea on click
            copyArea.on('click', function () {
                navigator.clipboard.writeText($(this).text());
            })
        });
    };

    function formatInfo(input, info, copyTarget) {

        // Find inputs for members
        const emptyCopyArea = copyTarget.text() === '',
            id = input.attr('id'),
            lineBreak = '\n';

        let position = '';

        // For some reason not all member id's are strings???
        if (typeof (id) === 'string') {

            position = parseInt(id.replace('member-', ''));
        }

        // Start with Name
        let sabreSymbol = '§',
            sabreFormatting = '-' + Object.values(info[position])[0] + sabreSymbol;

        // Loop through each value for that row
        $.each(info[position], function (key, value) {

            // If value contains the info we need
            if (value.includes("3DOCS/DB") || value.includes("3CTCE") || value.includes("3CTCM") || value.includes("3DOCO") || value.includes("3DOCS/P/")) {

                // Append value to formatted string
                emptyCopyArea ? sabreFormatting += value + sabreSymbol : sabreFormatting += value + '-' + crewIteration + '.1' + sabreSymbol;
            }
        });

        // If copy area is empty, append string as is, else append string with linebreak added
        emptyCopyArea ? copyTarget.append(sabreFormatting) : copyTarget.append(lineBreak + sabreFormatting);
    }

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

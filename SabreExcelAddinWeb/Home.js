(function ($) {
    "use strict";

    // TODO
    /*
        - Add Refresh Button
        - Make Select All add everything to copyarea
        - Add Styling
        - Add Toggle for Passport
        - Add Toggle for Meal Option:
        (AVMLA,BBMLA,BLMLA,CHMLA,DBMLA,FPMLA,GFMLA,HNMLA,KSMLA,LCMLA,LFMLA,LSMLA,MOMLA,NLMLA,NOMLA,RVMLA,SFMLA,SPMLA,VGMLA,VJMLA,VLMLA,VOMLA)
        - Add iteration toggle(Start with -2.1, -3.1, etc.)
    */

    // To account for more than 1 member per record
    let crewIteration = 1,
        textBox = [];

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
                    memberHtml = $(`<div class="selection selection-${m}"><input id="member-${m}" class="member" name="member-${m}" type="checkbox" /><label for="member-${m}">${memberName}</label><div class="optional-items"><label>Choose a Frequent Flyer:</label><select name="FF-select" id="FF-select-${m}"><option value="">-</option></select></div></div>`);

                $('.crew-member-select-container').append(memberHtml);

                // Loop through values
                for (let i = 0; i < value.length; i++) {

                    // If value contain Frequent Flyer ID
                    if (/^FF.*$/.test(value[i])) {

                        // If value contains multiple FF in one cell
                        if (/^FF.*FF.*$/.test(value[i])) {

                            // Split value by whitespace
                            const multipleFF = value[i].split(/[\s]/);

                            // Loop through array of split values
                            for (let j = 0; j < multipleFF.length; j++) {

                                // If value starts with "FF"
                                if (multipleFF[j].substring(0, 2) === "FF") {

                                    // Append each value to dropdown
                                    $(`#FF-select-${m}`).append($(`<option value="${multipleFF[j]}">${multipleFF[j]}</option>`));
                                }
                            }
                        } else {
                            // Append value to dropdown
                            $(`#FF-select-${m}`).append($(`<option value="${value[i]}">${value[i]}</option>`));
                        }
                    }
                }

                m++;
            });

            // When clicking member
            $('input.member').on('click', function () {
                const member = $(this);

                // If member is checked
                if (member.is(':checked')) {
                    // Add member to textbox and increase iteration
                    if (formatInfo(member, crewInfo, copyArea)) crewIteration++;

                // Else
                } else {
                    // Remove member from textbox and decrease iteration
                    if (removeInfo(member, copyArea)) crewIteration--;
                }
            });

            // When choosing Frequent Flyer
            $('select[name=FF-select]').on('change', function () {
                const dropdown = $(this),
                    ff = dropdown.val();

                addFrequentFlyer(dropdown, ff, copyArea);
            });

            // Select all inputs
            $('#selectAll').on('click', selectAll);

            // Copy text from textarea on click
            copyArea.on('click', function () {
                navigator.clipboard.writeText($(this).text());
            })
        });
    };

    function addFrequentFlyer(input, value, copyTarget) {

        const id = input.attr('id').replace('FF-select-', ''),
            checkBox = $(`input#member-${id}`),
            label = checkBox.siblings('label').text(),
            sabreSymbol = '§';

        // If checkbox isn't checked, check it
        if (checkBox.prop('checked') === false) {
            checkBox.click();
        }

        // Loop through textbox items
        for (let i = 0; i < textBox.length; i++) {

            // If item in textbox matches label name
            if (textBox[i].includes(label)) {

                // If dropdown option is blank
                if (value === "") {

                    // If textboxt item has iteration
                    if (textBox[i].includes('.1')) {

                        // Replace FF and iteration with empty string
                        textBox[i] = textBox[i].replace(/FF[a-zA-Z]*[0-9]*[-][0-9]*[\.]1§/, "");

                    // Else
                    } else {

                        // Replace FF with empty string
                        textBox[i] = textBox[i].replace(/FF[a-zA-Z]*[0-9]*§/, "");
                    }

                // Else
                } else {

                    // If items are iterated
                    if (crewIteration > 2) {

                        // Get iteration from textbox item
                        let memberIteration = textBox[i].trim().split(/.+(?=\-\d+\.1)/);

                        memberIteration = memberIteration[1].replace(sabreSymbol, "");

                        const replacementText = textBox[i] + value + memberIteration + sabreSymbol;

                        // Add Frequent flyer and iteration to text box
                        if (/FF[a-zA-Z]*[0-9]*§/.test(textBox[i])) {

                            const ffId = textBox[i].match(/FF[a-zA-Z]*[0-9]*§/)[0];
                            console.log(textBox[i].match(/FF[a-zA-Z]*[0-9]*§/)[0]);
                            console.log(textBox[i].replace(ffId, replacementText));

                            textBox[i] = textBox[i].replace(textBox[i].match(/FF[a-zA-Z]*[0-9]*§/)[0], replacementText);

                        } else if (/FF[a-zA-Z]*[0-9]*[-][0-9]*[\.]1§/.test(textBox[i])) {

                            const ffId = textBox[i].match(/FF[a-zA-Z]*[0-9]*[-][0-9]*[\.]1§/)[0];
                            console.log(textBox[i].match(/FF[a-zA-Z]*[0-9]*[-][0-9]*[\.]1§/)[0]);
                            console.log(textBox[i].replace(ffId, replacementText));

                            textBox[i] = textBox[i].replace(textBox[i].match(/FF[a-zA-Z]*[0-9]*[-][0-9]*[\.]1§/)[0], replacementText);

                        } else {

                            textBox[i] = textBox[i] + value + memberIteration + sabreSymbol;
                        }

                    // Else
                    } else {

                        const replacementText = textBox[i] = textBox[i] + value + sabreSymbol;

                        // Add Frequent Flyer to text box
                        if (/FF[a-zA-Z]*[0-9]*§/.test(textBox[i])) {

                            const ffId = textBox[i].match(/FF[a-zA-Z]*[0-9]*§/)[0];
                            console.log(textBox[i].match(/FF[a-zA-Z]*[0-9]*§/)[0]);
                            console.log(textBox[i].replace(ffId, replacementText));

                            textBox[i] = textBox[i].replace(textBox[i].match(/FF[a-zA-Z]*[0-9]*§/)[0], replacementText);

                        } else if (/FF[a-zA-Z]*[0-9]*[-][0-9]*[\.]1§/.test(textBox[i])) {

                            const ffId = textBox[i].match(/FF[a-zA-Z]*[0-9]*[-][0-9]*[\.]1§/)[0];
                            console.log(textBox[i].match(/FF[a-zA-Z]*[0-9]*[-][0-9]*[\.]1§/)[0]);
                            console.log(textBox[i].replace(ffId, replacementText));

                            textBox[i] = textBox[i].replace(textBox[i].match(/FF[a-zA-Z]*[0-9]*[-][0-9]*[\.]1§/)[0], replacementText);

                        } else {

                            textBox[i] = textBox[i] + value + sabreSymbol;
                        }
                    }
                }
            }
        }

        // Empty text area
        copyTarget.text("");

        // Add new text options
        copyTarget.append(textBox);
    }

    function formatInfo(input, info, copyTarget) {

        // Find inputs for members
        const emptyCopyArea = copyTarget.text() === '',
            id = input.attr('id'),
            lineBreak = '\n',
            sabreSymbol = '§';

        let position = '';

        // For some reason not all member id's are strings???
        if (typeof (id) === 'string') {

            position = parseInt(id.replace('member-', ''));
        }

        // Start with Name
        let sabreFormatting = '-' + Object.values(info[position])[0] + sabreSymbol;

        // Loop through each value for that row
        $.each(info[position], function (key, value) {

            // If value contains the info we need
            if (value.includes("3DOCS/DB") || value.includes("3CTCE") || value.includes("3CTCM") || value.includes("3DOCO") || value.includes("3DOCS/P/")) {

                // Append value to formatted string
                emptyCopyArea ? sabreFormatting += value + sabreSymbol : sabreFormatting += value + '-' + crewIteration + '.1' + sabreSymbol;
            }
        });

        // If copy area is empty, append string as is, else append string with linebreak added
        if (emptyCopyArea) {
            textBox.push(sabreFormatting);
            copyTarget.append(textBox);
        } else {
            textBox.push(lineBreak + sabreFormatting);
            copyTarget.text("");
            copyTarget.append(textBox);
        }

        return true;
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

    function removeInfo(input, copyTarget) {

        // Find info name
        const needsRemoval = input.siblings('label').text(),
            text = copyTarget.text();

        // If textbox includes name that's getting removed
        if (text.includes(needsRemoval)) {

            let newString = '',
                removedIteration = '';

            // Loop through textbox
            for (let i = 0; i < textBox.length; i++) {

                // If textbox item includes name
                if (textBox[i].includes(needsRemoval)) {

                    // Clear everything except -X.1
                    newString = textBox[i].replace(/.+(?=\-\d+\.1)/, '');
                    // If new string includes .1, trim string to number, else, set iteration to 0
                    newString.includes('.1') ? removedIteration = parseInt(newString.substring(0, newString.indexOf('1') + 1).trim().replace('-', '').replace('.1', '')) - 1 : removedIteration = 0;
                    // Remove string from array
                    textBox.splice(textBox.indexOf(textBox[i]), 1);
                }
            }

            // Loop through textbox again with new iteration
            for (let j = removedIteration; j < textBox.length; j++) {

                // If iteration starts at 0, clear first iteration
                if (j === 0) {
                    textBox[j] = textBox[j].replaceAll('-2.1', '').replace('\n', '');

                // Else replace text with correct iteration
                } else {
                    textBox[j] = textBox[j].replaceAll("-" + (j + 2) + ".1", "-" + (j + 1) + ".1");
                }
            };

            // Clear textbox
            copyTarget.text("");

            // Add new items
            copyTarget.append(textBox);

        }

        return true;
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

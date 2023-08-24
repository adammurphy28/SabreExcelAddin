(function ($) {
    "use strict";

    // TODO
    /*
        - Add Styling
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
                    memberHtml = $(`<div class="selection selection-${m}"><input id="member-${m}" class="member" name="member-${m}" type="checkbox" /><label for="member-${m}">${memberName}</label><div class="optional-items"></div></div>`);

                $('.crew-member-select-container').append(memberHtml);

                const optionalItems = $(`.selection-${m} .optional-items`);

                // Loop through values
                for (let i = 0; i < value.length; i++) {

                    // If value contain Frequent Flyer ID
                    if (/^FF.*$/.test(value[i])) {

                        // If container does not exist
                        if (!$(`#FF-select-${m}`).length > 0) {
                            optionalItems.append(`<div class="FF-container"><label>Choose a Frequent Flyer:</label><select name="FF-select" id="FF-select-${m}"><option value="">-</option></select></div>`)
                        }

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

                    // If value contains Passport
                    if (value[i].includes("3DOCS/P/")) {

                        // Remove any formatting inconsistencies
                        const valueFormatted = value[i].replace("/ ", "");

                        // Append Passport Container to Optional Items Container
                        optionalItems.append(`<div class="passport-container"><label for="passport-${m}">Choose a Passport:</label><select id="passport-select-${m}" class="passport-toggle" name="passport-select" type="checkbox"><option value="">-</option></select></div>`);

                        // If there are multiple Passports
                        if (/^(3DOCS\/P\/).*(3DOCS\/P\/).*$/.test(valueFormatted)) {

                            // Split Passports by space between each
                            const multiplePassports = valueFormatted.split(/(?<=[a-zA-Z]*[0-9]*)[\s](?=3DOCS\/P\/)/g);

                            // For each Passport
                            for (let j = 0; j < multiplePassports.length; j++) {

                                // Verify value is equal to Passport format
                                if (multiplePassports[j].substring(0, 8) === "3DOCS/P/") {

                                    // Add Passport as option
                                    $(`#passport-select-${m}`).append(`<option value="${multiplePassports[j]}">${multiplePassports[j].substring(0, 11)}...</option>`);
                                }
                            }
                        // Else
                        } else {

                            // Add Passport as option
                            $(`#passport-select-${m}`).append(`<option value="${valueFormatted}">${valueFormatted.substring(0, 11)}...</option>`);
                        }
                    }

                    // If value contains Meal Preference
                    if (/3[a-zA-z]{2}MLA/.test(value[i])) {

                        optionalItems.append(`<div class="meal-container"><input class="meal-preference" data-meal-preference="${value[i]}" id="meal-${m}" name="meal-${m}" type="checkbox" /><label for="meal-${m}">Include Meal Preference</label></div>`);
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

            // When choosing Meal Preference
            $('input.meal-preference').on('click', function () {
                const mealToggle = $(this),
                    mealPreference = mealToggle.data('meal-preference');

                addMealPreference(mealToggle, mealPreference, copyArea);
            });

            // When chooseing Passport
            $('select[name=passport-select]').on('change', function () {
                const dropdown = $(this),
                    passport = dropdown.val();

                addPassport(dropdown, passport, copyArea);
            });

            $('#reload').on('click', function (e) {
                e.preventDefault();

                location.reload();
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
            ffRegex = new RegExp(/[§]F{2}[a-zA-Z].*?(?=[§])./),
            ffIterRegex = new RegExp(/[§]F{2}[a-zA-Z].*?[-][0-9]+[\.]1[§]/),
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

                let ffId = textBox[i].match(ffRegex);

                // If dropdown option is blank
                if (value === "") {

                    // If textboxt item has iteration
                    if (ffIterRegex.test(textBox[i])) {

                        ffId = textBox[i].match(ffIterRegex);

                        // Replace FF and iteration with empty string
                        textBox[i] = textBox[i].replace(ffId[0], sabreSymbol);

                    // Else
                    } else {

                        // Replace FF with empty string
                        textBox[i] = textBox[i].replace(ffId[0], sabreSymbol);
                    }

                // Else
                } else {

                    let replacementText = sabreSymbol + value + sabreSymbol;

                    // If items are iterated
                    if (crewIteration > 2) {

                        // Get iteration from textbox item
                        let memberIteration = textBox[i].trim().split(/.+(?=\-\d+\.1)/);

                        memberIteration.length > 1 ? memberIteration = memberIteration[1].replace(sabreSymbol, "") : memberIteration = "";

                        replacementText = sabreSymbol + value + memberIteration + sabreSymbol;

                        // Add Frequent flyer and iteration to text box
                        if (ffRegex.test(textBox[i])) {

                            textBox[i] = textBox[i].replace(ffId[0], replacementText);

                        } else if (ffIterRegex.test(textBox[i])) {

                            ffId = textBox[i].match(ffIterRegex);

                            textBox[i] = textBox[i].replace(ffId[0], replacementText);

                        } else {

                            textBox[i] = textBox[i] + value + memberIteration + sabreSymbol;
                        }

                    // Else
                    } else {

                        // Add Frequent Flyer to text box
                        if (ffRegex.test(textBox[i])) {

                            textBox[i] = textBox[i].replace(ffId[0], replacementText);

                        } else if (ffIterRegex.test(textBox[i])) {

                            ffId = textBox[i].match(ffIterRegex);

                            textBox[i] = textBox[i].replace(ffId[0], replacementText);

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

    function addMealPreference(input, value, copyTarget) {
        const id = input.attr('id').replace('meal-', ''),
            checkBox = $(`input#member-${id}`),
            mpRegex = new RegExp(/[§]3[a-zA-z]{2}MLA[§]/),
            mpIterRegex = new RegExp(/[§]3[a-zA-z]{2}MLA[-][0-9]+[\.]1[§]/),
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

                let mpId = textBox[i].match(mpRegex);

                // If toggle is turned off
                if (input.prop('checked') === false) {

                    // If textboxt item has iteration
                    if (mpIterRegex.test(textBox[i])) {

                        mpId = textBox[i].match(mpIterRegex);

                        // Replace MP and iteration with empty string
                        textBox[i] = textBox[i].replace(mpId[0], sabreSymbol);

                    // Else
                    } else {

                        // Replace MP with empty string
                        textBox[i] = textBox[i].replace(mpId[0], sabreSymbol);
                    }

                // Else
                } else {

                    let replacementText = sabreSymbol + value + sabreSymbol;

                    // If items are iterated
                    if (crewIteration > 2) {

                        // Get iteration from textbox item
                        let memberIteration = textBox[i].trim().split(/.+(?=\-\d+\.1)/);

                        memberIteration.length > 1 ? memberIteration = memberIteration[1].replace(sabreSymbol, "") : memberIteration = "";

                        replacementText = sabreSymbol + value + memberIteration + sabreSymbol;

                        // Add Meal Preference and iteration to text box
                        if (mpRegex.test(textBox[i])) {

                            textBox[i] = textBox[i].replace(mpId[0], replacementText);

                        } else if (mpIterRegex.test(textBox[i])) {

                            mpId = textBox[i].match(mpIterRegex);

                            textBox[i] = textBox[i].replace(mpId[0], replacementText);

                        } else {

                            textBox[i] = textBox[i] + value + memberIteration + sabreSymbol;
                        }

                        // Else
                    } else {

                        // Add Meal Preference to text box
                        if (mpRegex.test(textBox[i])) {

                            textBox[i] = textBox[i].replace(mpId[0], replacementText);

                        } else if (mpIterRegex.test(textBox[i])) {

                            mpId = textBox[i].match(mpIterRegex);

                            textBox[i] = textBox[i].replace(mpId[0], replacementText);

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

    function addPassport(input, value, copyTarget) {

        const id = input.attr('id').replace('passport-select-', ''),
            checkBox = $(`input#member-${id}`),
            label = checkBox.siblings('label').text(),
            passportRegex = new RegExp(/[§]3DOCS\/P\/[a-zA-Z].*?(?=§)./),
            passportIterRegex = new RegExp(/[§]3DOCS\/P\/[a-zA-Z].*?[-][0-9]+[\.]1[§]/),
            sabreSymbol = '§';

        // If checkbox isn't checked, check it
        if (checkBox.prop('checked') === false) {
            checkBox.click();
        }

        // Loop through textbox items
        for (let i = 0; i < textBox.length; i++) {

            // If item in textbox matches label name
            if (textBox[i].includes(label)) {

                let passportId = textBox[i].match(passportRegex);

                // If dropdown option is blank
                if (value === "") {

                    // If textboxt item has iteration
                    if (passportIterRegex.test(textBox[i])) {

                        passportId = textBox[i].match(passportIterRegex);

                        // Replace Passport and iteration with empty string
                        textBox[i] = textBox[i].replace(passportId[0], sabreSymbol);

                        // Else
                    } else {

                        // Replace Passport with empty string
                        textBox[i] = textBox[i].replace(passportId[0], sabreSymbol);
                    }

                    // Else
                } else {

                    let replacementText = sabreSymbol + value + sabreSymbol;

                    // If items are iterated
                    if (crewIteration > 2) {

                        // Get iteration from textbox item
                        let memberIteration = textBox[i].trim().split(/.+(?=\-\d+\.1)/);

                        memberIteration.length > 1 ? memberIteration = memberIteration[1].replace(sabreSymbol, "") : memberIteration = "";

                        replacementText = sabreSymbol + value + memberIteration + sabreSymbol;

                        // Add Passport and iteration to text box
                        if (passportRegex.test(textBox[i])) {

                            textBox[i] = textBox[i].replace(passportId[0], replacementText);

                        } else if (passportIterRegex.test(textBox[i])) {

                            passportId = textBox[i].match(passportIterRegex);

                            textBox[i] = textBox[i].replace(passportId[0], replacementText);

                        } else {

                            textBox[i] = textBox[i] + value + memberIteration + sabreSymbol;
                        }

                        // Else
                    } else {

                        // Add Passport to text box
                        if (passportRegex.test(textBox[i])) {

                            textBox[i] = textBox[i].replace(passportId[0], replacementText);

                        } else if (passportIterRegex.test(textBox[i])) {

                            passportId = textBox[i].match(passportIterRegex);

                            textBox[i] = textBox[i].replace(passportId[0], replacementText);

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
            if (value.includes("3DOCS/DB") || value.includes("3CTCE") || value.includes("3CTCM") || value.includes("3DOCO")) {

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
        const id = input.attr('id').replace('member-', ''),
            needsRemoval = input.siblings('label').text(),
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

                    // Clear out FF dropdown
                    $(`#FF-select-${id}`).val('');

                    // Uncheck Meal Preference
                    $(`#meal-${id}`).prop('checked', false);

                    // Clear out Passport dropdown
                    $(`#passport-select-${id}`).val('');
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

        $('.member').each(function () {

            $(this).click();
        });
    }

}(this.jQuery));

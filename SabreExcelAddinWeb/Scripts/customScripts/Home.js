(function ($) {
    "use strict";

    const sabreSymbol = "§"

    // To account for more than 1 member per record
    let correctFormat = true,
        crewIteration = 1,
        textBox = [];

    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function () {

        $(document).ready(async function () {

            const copyArea = $('.copyArea'),
                crewInfo = await getCrewInfo();
            let m = 0;

            if (correctFormat) {

                // Loop through filtered items
                $.each(crewInfo, function (key, value) {

                    const memberName = value[0],
                        memberHtml = $(`<div class="selection selection-${m}"><div class="member-name"><input id="member-${m}" class="member" name="member-${m}" type="checkbox" /><label for="member-${m}">${memberName}</label></div></div>`);

                    $('.crew-member-select-container').append(memberHtml);

                    const currentSelection = $(`.selection-${m}`),
                        containsOptional = $(`.selection-${m} .optional-items`),
                        optionalItems = $(`<input class="optional-items-dropdown" id="optional-items-dropdown-${m}" name="optional-items-dropdown-${m}" type="checkbox" /><label for="optional-items-dropdown-${m}"><i class="fa-solid fa-angle-up"></i></label><div class="optional-items"></div>`);

                    // Loop through values
                    for (let i = 0; i < value.length; i++) {

                        // If value contain Frequent Flyer ID
                        if (/^FF.*$/.test(value[i])) {

                            // If selection member does not contain optional items dropdown
                            if (!containsOptional.length > 0) {
                                currentSelection.append(optionalItems);
                            }

                            // If container does not exist
                            if (!$(`#FF-select-${m}`).length > 0) {
                                $(`.selection-${m} .optional-items`).append(`<div class="FF-container"><label>Choose a Frequent Flyer:</label><select name="FF-select" id="FF-select-${m}"><option value="">-</option></select></div>`)
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

                            // If selection member does not contain optional items dropdown
                            if (!containsOptional.length > 0) {
                                currentSelection.append(optionalItems);
                            }

                            // Append Passport Container to Optional Items Container
                            $(`.selection-${m} .optional-items`).append(`<div class="passport-container"><label for="passport-${m}">Choose a Passport:</label><select id="passport-select-${m}" class="passport-toggle" name="passport-select" type="checkbox"><option value="">-</option></select></div>`);

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

                            // If selection member does not contain optional items dropdown
                            if (!containsOptional.length > 0) {
                                currentSelection.append(optionalItems);
                            }

                            $(`.selection-${m} .optional-items`).append(`<div class="meal-container"><input class="meal-preference" data-meal-preference="${value[i]}" id="meal-${m}" name="meal-${m}" type="checkbox" /><label for="meal-${m}">Include Meal Preference</label></div>`);
                        }
                    }

                    m++;
                });

            } else {
                
                $('<p class="error">Incorrect Formatting. Please add one column with "Name" (ex. LastName/FirstName MiddleName) and one column with birthday(ex. 3DOCS/DB/01JAN2023/G/LASTNAME/FIRSTNAME/MIDDLENAME)(Please substitute "G" in birthday for gender initial. M or F)</p>').insertBefore($('.copyArea'));
            }

            // When clicking member
            $('input.member').on('click', function () {
                const member = $(this),
                    optionalItems = member.parents('.selection').children('.optional-items-dropdown');

                // If member is checked
                if (member.is(':checked')) {
                    // Add member to textbox and increase iteration
                    if (formatInfo(member, crewInfo, copyArea)) crewIteration++;

                    if (!optionalItems.is(':checked')) {
                        optionalItems.click();
                    }

                // Else
                } else {
                    // Remove member from textbox and decrease iteration
                    if (removeInfo(member, copyArea)) crewIteration--;

                    if (optionalItems.is(':checked')) {
                        optionalItems.click();
                    }
                }
            });

            // When changing iteration
            $('select[name=iteration-value]').on('change', function () {
                const dropdown = $(this),
                    iterationValue = dropdown.val();

                changeStartingIteration(iterationValue, copyArea);
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

            // Reload Button functionality
            $('#reload').on('click', function (e) {
                e.preventDefault();

                location.reload();
            });

            // Deselect all inputs
            $('#deselectAll').on('click', deselectAll);

            // Select all inputs
            $('#selectAll').on('click', selectAll);

            //Scroll to Bottom
            $('#toBottom').on('click', scrollToBottom);

            // Scroll to Top
            $('#toTop').on('click', scrollToTop);

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
            label = checkBox.siblings('label').text();

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
            label = checkBox.siblings('label').text();        

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
            passportIterRegex = new RegExp(/[§]3DOCS\/P\/[a-zA-Z].*?[-][0-9]+[\.]1[§]/);

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

    function changeStartingIteration(value, copyTarget) {

        // Regex for -X.1
        const iterationRegex = new RegExp(/[-][0-9]+[\.]1[§]/g);

        // New iteration count
        let newIteration = value;

        // If value is empty
        if (value === "") {

            // Set iteration to 1
            newIteration = 1;
        }

        // If textarea is empty, set iteration to current value
        if (copyTarget.text() === '') {

            crewIteration = value;

        } else {

            // Loop through textbox items
            for (let i = 0; i < textBox.length; i++) {

                // Setting format with new iteration
                const newFormat = `-${newIteration}.1${sabreSymbol}`;

                // If first item in loop
                if (i === 0) {

                    // If iteration count is equal to 1
                    if (newIteration === 1) {

                        // Replace any previous iterations with nothing
                        textBox[i] = textBox[i].replaceAll(iterationRegex, sabreSymbol).replace("\n", "");

                    // Else
                    } else {

                        // Split string starting with the name
                        let name = textBox[i].trim().split(sabreSymbol, 1),
                            newString = textBox[i].replace(name + sabreSymbol, "");

                        // Test string for regex and replace all iteration symbols if matches
                        iterationRegex.test(newString) ? newString = newString.replaceAll(iterationRegex, newFormat).replace("\n", "") : newString = newString.replaceAll(sabreSymbol, newFormat).replace("\n", "");

                        // Add new string to array
                        textBox[i] = name + sabreSymbol + newString;
                    }

                // Else
                } else {

                    // Replace all iterations with new format
                    textBox[i] = textBox[i].replaceAll(iterationRegex, newFormat);

                }

                // Increase iteration
                newIteration++;
            }

            // Replace crewIteration with new iteration
            crewIteration = newIteration;
        }

        // Empty text area
        copyTarget.text("");

        // Add new text options
        copyTarget.append(textBox);
    }

    function deselectAll(event) {
        event.preventDefault();

        if ($('#selectAll').prop('checked') === true) {

            $('#selectAll').click();
        } else {

            // Loop through items
            $('.member').each(function () {

                const member = $(this);

                // If item is checked
                if (member.prop('checked') === true) {

                    // Click to uncheck
                    member.click();
                }
            });
        }
    }

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
        let sabreFormatting = '-' + Object.values(info[position])[0] + sabreSymbol;

        // Loop through each value for that row
        $.each(info[position], function (key, value) {

            // If value contains the info we need
            if (value.includes("3DOCS/DB") || value.includes("3CTCE") || value.includes("3CTCM") || value.includes("3DOCO")) {

                // Append value to formatted string
                emptyCopyArea && crewIteration === 1 ? sabreFormatting += value + sabreSymbol : sabreFormatting += value + '-' + crewIteration + '.1' + sabreSymbol;
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
            let i = 0,
                nameColumn = 0;

            // Loop through each row
            $.each(cellsWithText, function (rowKey, rowValue) {

                let values = [];

                // Loop through each cell in the row
                $.each(rowValue, function (columnKey, columnValue) {

                    // If it's the first row and column value is equal to name
                    if (rowKey === 0 && columnValue.toUpperCase() === "NAME") {

                        // Update Name Column Number
                        nameColumn = columnKey;

                    }

                    // If cell is not empty
                    if (columnValue !== "" && columnValue !== null && columnValue !== '') {

                        // Transform everything to uppercase
                        let valueFormatted = columnValue.toUpperCase();

                        // Remove spaces at beginning of cell
                        if (valueFormatted.charAt(0) === " ") {
                            valueFormatted = valueFormatted.substring(1, valueFormatted.length);
                        }

                        // Remove spaces at end of cell
                        if (valueFormatted.charAt(valueFormatted.length) === " ") {
                            valueFormatted = valueFormatted.substring(0, valueFormatted.length - 1);
                        }

                        // If current column is equal to name column, add value to beginning of array, else add to end
                        columnKey === nameColumn ? values.unshift(valueFormatted) : values.push(valueFormatted);
                    }
                });

                // Remove any section headers
                if (rowKey === 0 && !values.includes("NAME")) {
                    correctFormat = false;

                    return false;
                } else if (values.length > 1 && !values.includes("NAME")) {
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

        const sortedItems = Object.values(filteredItems).sort();

        filteredItems = sortedItems;

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
            iterationRegex = new RegExp(/[-][0-9]+[\.]1[§]/g),
            needsRemoval = input.siblings('label').text(),
            text = copyTarget.text();

        // If textbox includes name that's getting removed
        if (text.includes(needsRemoval)) {

            let newString = '',
                removedIteration = '';

            // Loop through textbox to get current iteration from where item is removed
            for (let i = 0; i < textBox.length; i++) {

                // If textbox item includes name
                if (textBox[i].includes(needsRemoval)) {

                    // Clear everything except -X.1
                    newString = textBox[i].replace(/.+(?=\-\d+\.1)/, '');
                    // If new string includes .1, trim string to number, else, set iteration to 0
                    newString.includes('.1') ? removedIteration = parseInt(newString.substring(0, newString.indexOf('1') + 1).trim().replace('-', '').replace('.1', '')) : removedIteration = 2;
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

            let iteration = removedIteration;

            // Loop through textbox again with new iteration
            for (let j = 0; j < textBox.length; j++) {

                // Set new iteration format
                let newIteration = "-" + iteration + ".1" + sabreSymbol;

                // If old iteration is -2.1
                if (iteration === 2) {

                    // Replace old iteration with empty string
                    textBox[j] = textBox[j].replaceAll('-2.1' + sabreSymbol, sabreSymbol).replace('\n', '');

                    // Increase iteration if successful
                    iteration++;

                // Else replace text with correct iteration
                }
                else {

                    // Replace everything that isn't -X.1 with empty string then parse for "X"
                    let currentString = textBox[j].replace(/.+(?=\-\d+\.1)/, ''),
                        testNumber = parseInt(currentString.substring(0, currentString.indexOf('1') + 1).trim().replace('-', '').replace('.1', ''));

                    // If current iteration is less than number above
                    if (iteration < testNumber) {

                        // Replace previous iteration with new iteration
                        textBox[j] = textBox[j].replaceAll(iterationRegex, newIteration);

                        // Increase iteration if successful
                        iteration++;
                    }
                }
            };

            // Clear textbox
            copyTarget.text("");

            // Add new items
            copyTarget.append(textBox);

        }

        return true;
    }

    function scrollToBottom(event) {
        event.preventDefault();

        const height = $('body').height();

        $('html, body').animate({ scrollTop: height }, "fast");
    }

    function scrollToTop(event) {
        event.preventDefault();

        $('html, body').animate({ scrollTop: 0 }, "fast");
    }

    function selectAll() {

        $('.member').each(function () {

            $(this).click();
        });
    }

}(this.jQuery));

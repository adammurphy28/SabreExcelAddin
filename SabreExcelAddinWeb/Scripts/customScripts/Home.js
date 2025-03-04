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
                crewInfo = await getCrewInfo(),
                sortedItems = Object.values(crewInfo).sort();

            if (correctFormat) {
                displayInfo(sortedItems);
            }
            else {
                
                $('<p class="error">Incorrect Formatting. Please add one column with "Name" (ex. LastName/FirstName MiddleName) and one column with birthday(ex. 3DOCS/DB/01JAN2023/G/LASTNAME/FIRSTNAME/MIDDLENAME)(Please substitute "G" in birthday for gender initial. M or F)</p>').insertBefore($('.copyArea'));
            }

            // When disabling alphabetical order
            $('#disable-order').on('change', function () {

                const checkbox = $(this);

                let order = false;

                if (checkbox.is(':checked')) {

                    clearMembers();

                    $('.letter-nav').remove();

                    $('body').addClass('no-letter-nav');

                    reset();

                    displayInfo(crewInfo, order);

                } else {

                    order = true;

                    clearMembers();

                    $('body').removeClass('no-letter-nav');

                    $('<div class="letter-nav"></div>').insertAfter(copyArea);

                    reset();

                    displayInfo(sortedItems, order);
                }

            });

            // When clicking member
            $('.crew-member-select-container').on('change', 'input.member', function () {
                const member = $(this),
                    optionalItems = member.parents('.selection').children('.optional-items-dropdown');

                let infoData = sortedItems;

                // If alphabetical order is turned off
                if ($('#disable-order').is(':checked')) {
                    infoData = crewInfo;
                }

                // If member is checked
                if (member.is(':checked')) {

                    // Add member to textbox and increase iteration
                    if (formatInfo(member, infoData, copyArea)) crewIteration++;

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
            $('.crew-member-select-container').on('change', 'select[name=iteration-value]', function () {
                const dropdown = $(this),
                    iterationValue = dropdown.val();

                changeStartingIteration(iterationValue, copyArea);
            });

            // When choosing Frequent Flyer
            $('.crew-member-select-container').on('change', 'select[name=FF-select]', function () {
                const dropdown = $(this),
                    ff = dropdown.val(),
                    airline = ff.substring(0, 4).replace('FF', ''),
                    americanAirlinePartners = [
                        {
                            partner: 'Aer Lingus',
                            code: 'EI'
                        },
                        {
                            partner: 'Air Tahiti Nui',
                            code: 'TN'
                        },
                        {
                            partner: 'Alaska Airlines',
                            code: 'AS'
                        },
                        {
                            partner: 'American Airlines',
                            code: 'AA'
                        },
                        {
                            partner: 'British Airways',
                            code: 'BA'
                        },
                        {
                            partner: 'Cape Air',
                            code: '9K'
                        },
                        {
                            partner: 'Cathay Pacific',
                            code: 'CX'
                        },
                        {
                            partner: 'China Southern Airlines',
                            code: 'CZ'
                        },
                        {
                            partner: 'Etihad Airways',
                            code: 'EY'
                        },
                        {
                            partner: 'Fiji Airways',
                            code: 'FJ'
                        },
                        {
                            partner: 'Finnair',
                            code: 'AY'
                        },
                        {
                            partner: 'GOL Airlines',
                            code: 'G3'
                        },
                        {
                            partner: 'Hawaiian Airlines',
                            code: 'HA'
                        },
                        {
                            partner: 'Iberia',
                            code: 'IB'
                        },
                        {
                            partner: 'Japan Airlines',
                            code: 'JL'
                        },
                        {
                            partner: 'JetBlue*',
                            code: 'B6'
                        },
                        {
                            partner: 'Malaysia Airlines',
                            code: 'MH'
                        },
                        {
                            partner: 'Qantas',
                            code: 'QF'
                        },
                        {
                            partner: 'Qatar Airways',
                            code: 'QR'
                        },
                        {
                            partner: 'Royal Air Maroc',
                            code: 'AT'
                        },
                        {
                            partner: 'Royal Jordanian Airlines',
                            code: 'RJ'
                        },
                        {
                            partner: 'Silver Airways',
                            code: '3M'
                        },
                        {
                            partner: 'SriLankan Airlines',
                            code: 'UL'
                        },
                    ],
                    deltaAirlinePartners = [
                        {
                            partner: 'Aerolineas Argentinas',
                            code: 'AR'
                        },
                        {
                            partner: 'Aeromexico',
                            code: 'AM'
                        },
                        {
                            partner: 'Air Europa',
                            code: 'UX'
                        },
                        {
                            partner: 'Air france',
                            code: 'AF'
                        },
                        {
                            partner: 'Cape Air',
                            code: '9K'
                        },
                        {
                            partner: 'China Airlines',
                            code: 'CI'
                        },
                        {
                            partner: 'China eastern',
                            code: 'MU'
                        },
                        {
                            partner: 'Delta Airlines',
                            code: 'DL'
                        },
                        {
                            partner: 'Garuda Indonesia',
                            code: 'GA'
                        },
                        {
                            partner: 'Hawaiian Airlines',
                            code: 'HA'
                        },
                        {
                            partner: 'Kenya Airways',
                            code: 'KQ'
                        },
                        {
                            partner: 'KLM',
                            code: 'KL'
                        },
                        {
                            partner: 'Korean air',
                            code: 'KE'
                        },
                        {
                            partner: 'Latam',
                            code: 'LA'
                        },
                        {
                            partner: 'Middle East Airlines',
                            code: 'ME'
                        },
                        {
                            partner: 'Scandinavian',
                            code: 'SK'
                        },
                        {
                            partner: 'Saudia',
                            code: 'SV'
                        },
                        {
                            partner: 'Tarom',
                            code: 'RO'
                        },
                        {
                            partner: 'Vietnam Airlines',
                            code: 'VN'
                        },
                        {
                            partner: 'Virgin atlantic',
                            code: 'VS'
                        },
                        {
                            partner: 'West jet',
                            code: 'WS'
                        },
                        {
                            partner: 'XiamenAir',
                            code: 'MF'
                        },
                    ],
                    scandinavianAirlinePartners = [
                        {
                            partner: 'Aerolineas Argentinas',
                            code: 'AR'
                        },
                        {
                            partner: 'Aeromexico',
                            code: 'AM'
                        },
                        {
                            partner: 'Air Europa',
                            code: 'UX'
                        },
                        {
                            partner: 'Air france',
                            code: 'AF'
                        },
                        {
                            partner: 'China Airlines',
                            code: 'CI'
                        },
                        {
                            partner: 'China eastern',
                            code: 'MU'
                        },
                        {
                            partner: 'Delta Airlines',
                            code: 'DL'
                        },
                        {
                            partner: 'Garuda Indonesia',
                            code: 'GA'
                        },
                        {
                            partner: 'KLM',
                            code: 'KL'
                        },
                        {
                            partner: 'Korean air',
                            code: 'KE'
                        },
                        {
                            partner: 'Middle East Airlines',
                            code: 'ME'
                        },
                        {
                            partner: 'Scandinavian',
                            code: 'SK'
                        },
                        {
                            partner: 'Saudia',
                            code: 'SV'
                        },
                        {
                            partner: 'Tarom',
                            code: 'RO'
                        },
                        {
                            partner: 'Vietnam Airlines',
                            code: 'VN'
                        },
                        {
                            partner: 'Virgin atlantic',
                            code: 'VS'
                        },
                        {
                            partner: 'XiamenAir',
                            code: 'MF'
                        },
                    ],
                    unitedAirlinePartners = [
                        {
                            partner: 'Aegean Airlines',
                            code: 'A3'
                        },
                        {
                            partner: 'Air Canada',
                            code: 'AC'
                        },
                        {
                            partner: 'Air China',
                            code: 'CA'
                        },
                        {
                            partner: 'Air Dolomiti',
                            code: 'EN'
                        },
                        {
                            partner: 'Air India',
                            code: 'AI'
                        },
                        {
                            partner: 'Air New Zealand',
                            code: 'NZ'
                        },
                        {
                            partner: 'All Nippon (ANA)',
                            code: 'NH'
                        },
                        {
                            partner: 'Asiana Airlines',
                            code: 'OZ'
                        },
                        {
                            partner: 'Austrian Airlines',
                            code: 'OS'
                        },
                        {
                            partner: 'Avianca',
                            code: 'AV'
                        },
                        {
                            partner: 'Azul',
                            code: 'AD'
                        },
                        {
                            partner: 'Brussels Airlines',
                            code: 'SN'
                        },
                        {
                            partner: 'Cape Air',
                            code: '9K'
                        },
                        {
                            partner: 'Copa Airlines',
                            code: 'CM'
                        },
                        {
                            partner: 'Croatia Airlines',
                            code: 'OU'
                        },
                        {
                            partner: 'Edelweiss Air',
                            code: 'WK'
                        },
                        {
                            partner: 'Egypt Air',
                            code: 'MS'
                        },
                        {
                            partner: 'Ethiopian Airlines',
                            code: 'ET'
                        },
                        {
                            partner: 'Eurowings',
                            code: 'EW'
                        },
                        {
                            partner: 'Eva Airways',
                            code: 'BR'
                        },
                        {
                            partner: 'Hawaiian Airlines',
                            code: 'HA'
                        },
                        {
                            partner: 'ITA Airways',
                            code: 'AZ'
                        },
                        {
                            partner: 'LOT Polish Airlines',
                            code: 'LO'
                        },
                        {
                            partner: 'Lufthansa',
                            code: 'LH'
                        },
                        {
                            partner: 'Shenzhen Airlines',
                            code: 'ZH'
                        },
                        {
                            partner: 'Silver Airways',
                            code: '3M'
                        },
                        {
                            partner: 'Singapore Airlines',
                            code: 'SQ'
                        },
                        {
                            partner: 'Swiss Airlines',
                            code: 'LX'
                        },
                        {
                            partner: 'TAP Air',
                            code: 'TP'
                        },
                        {
                            partner: 'Thai Airways',
                            code: 'TG'
                        },
                        {
                            partner: 'Turkish Airlines',
                            code: 'TK'
                        },
                        {
                            partner: 'United Airlines',
                            code: 'UA'
                        },
                    ],
                    partnerContainer = dropdown.parent().siblings('.partner-container').find('select[name="partner-select"]');

                addFrequentFlyer(dropdown, ff, copyArea);

                // Function to add partners to dropdown based on FF
                function addPartners(partners) { 

                    // Loop through partner variable passed
                    partners.forEach(function (p) {
                        // Find the partner container relevant to this person and append the airline partner info to the dropdown
                        if (p.code !== airline) {
                            partnerContainer.append(`<option value="${p.code}">${p.partner}</option>`);
                        }
                    });
                }

                // Check Airline partners to see if selected FF has partners
                (function () {

                    // Clear HTML beforehand
                    partnerContainer.html('<option value="">-</option>');

                    americanAirlinePartners.forEach(function (aa) {
                        if (aa.code == airline) {
                            addPartners(americanAirlinePartners);
                            return;
                        }
                    });

                    deltaAirlinePartners.forEach(function (dl) {
                        if (dl.code == airline && sk.code !== 'SK') {
                            addPartners(deltaAirlinePartners);
                            return;
                        }
                    });

                    scandinavianAirlinePartners.forEach(function (sk) {
                        if (sk.code == airline && sk.code !== 'DL') {
                            addPartners(scandinavianAirlinePartners);
                            return;
                        }
                    });

                    unitedAirlinePartners.forEach(function (ua) {
                        if (ua.code == airline) {
                            addPartners(unitedAirlinePartners);
                            return;
                        }
                    });
                }());
            });

            // When choosing Meal Preference
            $('.crew-member-select-container').on('click', 'input.meal-preference', function () {
                const mealToggle = $(this),
                    mealPreference = mealToggle.data('meal-preference');

                addMealPreference(mealToggle, mealPreference, copyArea);
            });

            // When choosing Partner Preference
            $('.crew-member-select-container').on('change', 'select[name=partner-select]', function () {
                const dropdown = $(this),
                    partner = dropdown.val();

                addFrequentFlyer(dropdown, partner, copyArea, true);
            });

            // When chooseing Passport
            $('.crew-member-select-container').on('change', 'select[name=passport-select]', function () {
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
    function addFrequentFlyer(input, value, copyTarget, hasPartner = false) {

        const id = input.attr('id').replace(/^.*-select-/, ''),
            checkBox = $(`input#member-${id}`),
            ffNumber = new RegExp(/(?<=§)F{2}.*?(?=§)/),
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

                let ffId = textBox[i].match(ffRegex),
                    // Get only the FF number
                    number = hasPartner ? textBox[i].match(ffNumber) : '';

                if (number !== null && number !== '') {
                    // Remove anything leftover other than the number
                    number = number[0].split('/');
                    number = number[0].split('-');
                }

                // If dropdown option is blank
                if (value === "") {

                    // If we're deadline with a partner airline
                    if (hasPartner) {

                        // Grab the partner identifier
                        let airlinePartner = new RegExp(`(?<=${number[0]})\/..`);

                        // Find that match
                        ffId = textBox[i].match(airlinePartner);

                        // Replace with empty string
                        textBox[i] = textBox[i].replace(ffId[0], '');
                    }
                    // If textboxt item has iteration
                    else if (ffIterRegex.test(textBox[i])) {

                        ffId = textBox[i].match(ffIterRegex);

                        // Replace FF and iteration with empty string
                        textBox[i] = textBox[i].replace(ffId[0], sabreSymbol);

                    }
                    // Else
                    else {

                        // Replace FF with empty string
                        textBox[i] = textBox[i].replace(ffId[0], sabreSymbol);
                    }

                // Else
                } else {

                    // Check for partner airline and handle accordingly
                    let replacementText = hasPartner ? sabreSymbol + number[0] + '/' + value + sabreSymbol : sabreSymbol + value + sabreSymbol;

                    // If items are iterated
                    if (crewIteration > 2) {

                        // Get iteration from textbox item
                        let memberIteration = textBox[i].trim().split(/.+(?=\-\d+\.1)/);

                        memberIteration.length > 1 ? memberIteration = memberIteration[1].replace(sabreSymbol, "") : memberIteration = "";

                        if (hasPartner) {
                            replacementText = sabreSymbol + number[0] + '/' + value + memberIteration + sabreSymbol;
                        } else {
                            replacementText = sabreSymbol + value + memberIteration + sabreSymbol;
                        }

                        // Add Frequent flyer and iteration to text box
                        if (ffRegex.test(textBox[i])) {

                            textBox[i] = textBox[i].replace(ffId[0], replacementText);

                        } else if (ffIterRegex.test(textBox[i])) {

                            ffId = textBox[i].match(ffIterRegex);

                            textBox[i] = textBox[i].replace(ffId[0], replacementText);

                        } else {

                            if (hasPartner) {
                                textBox[i] = textBox[i] + number[0] + '/' + value + memberIteration + sabreSymbol;
                            } else {
                                textBox[i] = textBox[i] + value + memberIteration + sabreSymbol;
                            }
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

                            if (hasPartner) {
                                textBox[i] = textBox[i] + number[0] + '/' + value + sabreSymbol;
                            } else {
                                textBox[i] = textBox[i] + value + sabreSymbol;
                            }
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

    function clearMembers() {

        $('.selection').each(function () {

            const selection = $(this);

            if (selection.attr('id') !== 'first-selection') {

                selection.remove();
            }
        });

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

    function displayInfo(info, hasOrder = true) {

        let m = 0;

        // Loop through filtered items
        $.each(info, function (key, value) {

            const memberName = value[0],
                memberHtml = $(`<div class="selection selection-${m}"><div class="member-name"><input id="member-${m}" class="member" name="member-${m}" type="checkbox" /><label for="member-${m}">${memberName}</label></div></div>`);

            $('.crew-member-select-container').append(memberHtml);

            const currentSelection = $(`.selection-${m}`),
                containsOptional = $(`.selection-${m} .optional-items`),
                firstChar = memberName.charAt(0),
                letterNav = $('.letter-nav'),
                optionalItems = $(`<input class="optional-items-dropdown" id="optional-items-dropdown-${m}" name="optional-items-dropdown-${m}" type="checkbox" /><label for="optional-items-dropdown-${m}"><i class="fa-solid fa-angle-up"></i></label><div class="optional-items"></div>`);

            // Add Navigation by letters

            if (hasOrder) {

                if (!letterNav.find('a[href="#' + firstChar + '"]').length > 0) {

                    currentSelection.attr('id', firstChar);
                    letterNav.append($(`<a class="letter" href="#${firstChar}">${firstChar}</a>`));
                }
            }

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

                    // Add partner container if it doesn't exist
                    if (!$(`#partner-select-${m}`).length > 0) {
                        $(`<div class="partner-container"><label>Choose an Airline Partner:</label><select name="partner-select" id="partner-select-${m}"><option value="">-</option></select></div>`).insertAfter(`.selection-${m} .optional-items .FF-container`);
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

    function reset() {

        $('.copyArea').text('');
        $('#iteration-value').val('');
        textBox = [];
        crewIteration = 1;
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

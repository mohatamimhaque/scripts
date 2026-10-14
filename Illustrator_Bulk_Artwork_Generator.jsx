/**
 * =========================================================================================
 * UNIVERSAL BULK ARTWORK GENERATOR FOR ADOBE ILLUSTRATOR
 * =========================================================================================
 * File        : Illustrator_Bulk_Artwork_Generator.jsx
 * Engine      : Adobe ExtendScript (JavaScript 1.5 / ECMAScript 3)
 * Target App  : Adobe Illustrator CS6, CC 2014 – CC 2025+ (Windows & macOS)
 * Architecture: Data-Driven Variable Data Merge & Multi-Artboard Batch Cloner
 * Author      : Automated Artwork Automation Suite
 * =========================================================================================
 *
 * OVERVIEW
 * -----------
 * This ExtendScript automates high-volume, variable-data publishing in Adobe Illustrator.
 * It merges a structured JSON dataset with a pre-designed Illustrator (.ai) template,
 * generating a separate, fully populated artboard for every record in the dataset.
 *
 * Typical use cases include:
 *   - Event badges, conference passes, name tags, VIP credentials
 *   - Certificates of completion, awards, diplomas
 *   - Business cards, personalized invitations, tickets with serial numbers
 *   - Product labels, barcodes/SKU tags, customized marketing collateral
 *
 * -----------------------------------------------------------------------------------------
 * QUICK-START EXECUTION RUNBOOK
 * -----------------------------------------------------------------------------------------
 * 1. PREPARE DATASET (JSON):
 *    - Must be a JSON array of objects (UTF-8 encoded).
 *    - Example:
 *        [
 *          {"id": "001", "name": "Ada Lovelace", "role": "Speaker", "table": "A1"},
 *          {"id": "002", "name": "Alan Turing",  "role": "Keynote", "table": "B3"}
 *        ]
 *    - Tip: Use the companion Jupyter Notebook ('CSV_XLSX_to_Illustrator_JSON.ipynb')
 *      to convert any Excel (.xlsx/.xls) or CSV spreadsheet into this JSON format!
 *
 * 2. PREPARE ILLUSTRATOR TEMPLATE (.ai):
 *    - Create your single master design on Artboard 1.
 *    - In any text frame, insert placeholders matching your JSON keys wrapped in delimiters:
 *        e.g., {{id}}, {{name}}, {{role}}, {{table}}
 *    - Format fonts, font sizes, colors, tracking, and paragraph alignments as desired.
 *
 * 3. RUN SCRIPT IN ADOBE ILLUSTRATOR:
 *    - Open Adobe Illustrator (no document needs to be open beforehand).
 *    - Navigate to: File > Scripts > Other Script... (Shortcut: Ctrl+F12 on Windows / Cmd+F12 on macOS).
 *    - Select this 'Illustrator_Bulk_Artwork_Generator.jsx' file.
 *
 * 4. INTERACTIVE PROMPTS:
 *    - Prompt 1: Choose your JSON data file (*.json).
 *    - Prompt 2: Choose your master Illustrator template (*.ai).
 *    - Prompt 3: Confirm or enter the output base file name (without .ai extension).
 *
 * 5. AUTOMATIC EXECUTION:
 *    - The script scans the template and reports all detected placeholders.
 *    - Artboards are cloned into a grid layout according to template aspect ratio:
 *        * Landscape designs (width > height) -> 1 column vertical stack.
 *        * Portrait designs (width <= height) -> 10 column grid.
 *    - Artwork is duplicated, translated to center within each artboard, and text is dynamically replaced.
 *    - To prevent Illustrator memory crashes on large datasets, jobs are automatically
 *      partitioned into separate multi-part files of 100 artboards each:
 *        e.g., 'Badges-output-part-01.ai', 'Badges-output-part-02.ai', etc.
 *
 * -----------------------------------------------------------------------------------------
 * PLACEHOLDER MATCHING & FALLBACK RULES
 * -----------------------------------------------------------------------------------------
 * - Delimiters: Configurable via CONFIG (default is '{{' and '}}').
 * - Exact Match Priority:
 *     The script first attempts an exact, case-sensitive property lookup:
 *     {{name}} -> record["name"]
 * - Fuzzy Match Fallback:
 *     If an exact match is not found, both the placeholder and the JSON keys are
 *     normalized by converting to lowercase and stripping all non-alphanumeric characters.
 *     Example: {{Student ID}} will successfully resolve:
 *       - record["Student ID"]
 *       - record["student_id"]
 *       - record["studentId"]
 *       - record["STUDENTID"]
 * - Missing Field Handling:
 *     Controlled by CONFIG.missingFieldBehavior:
 *       - "blank" : (Default) Replaces unresolved placeholders with an empty string ("").
 *       - "keep"  : Preserves the placeholder tag intact in the text (e.g. "{{missing_key}}").
 *
 * -----------------------------------------------------------------------------------------
 * ARTBOARD GRID & LAYER ARCHITECTURE
 * -----------------------------------------------------------------------------------------
 * - Automatic Sizing: Uses the exact dimensions of Artboard 1 from the template.
 * - Dynamic Spacing : Configured via CONFIG.horizontalGap (50pt) and CONFIG.verticalGap (50pt).
 * - Artboard Naming : Automatically named after the first non-empty placeholder value found
 *                     (e.g., recipient's name), sanitized for valid filesystem characters.
 *                     Falls back to 'Item 001', 'Item 002', etc.
 * - Layer Isolation : When CONFIG.separateLayers is true, each record is placed on its own
 *                     dedicated layer (e.g. 'Item 001', 'Item 002') for easy editing.
 *
 * =========================================================================================
 */

/**
 * Global Engine Configuration
 * @property {string}  placeholderStart     - Opening delimiter string (default: "{{")
 * @property {string}  placeholderEnd       - Closing delimiter string (default: "}}")
 * @property {number}  columns              - Fallback column count for layout grids (default: 4)
 * @property {number}  horizontalGap        - Horizontal spacing between artboards in points (default: 50)
 * @property {number}  verticalGap          - Vertical spacing between artboards in points (default: 50)
 * @property {boolean} separateLayers      - Place each generated record on a dedicated layer (default: true)
 * @property {string}  missingFieldBehavior - Handling for missing keys: "blank" (clear) or "keep" (preserve tag)
 */
var CONFIG = {
    placeholderStart: "{{",
    placeholderEnd: "}}",

    columns: 4,

    horizontalGap: 50,
    verticalGap: 50,

    separateLayers: true,
    missingFieldBehavior: "blank"
};


/**
 * Main Controller Workflow.
 * Prompts user for JSON dataset and AI template, validates record array,
 * calculates batch chunking (100 items per file), detects placeholders,
 * duplicates layout grids, replaces dynamic text, and saves final AI files.
 */
function main() {
    var jsonFile = File.openDialog(
        "Select JSON file",
        "*.json"
    );

    if (!jsonFile) {
        return;
    }


    var records = readJSON(jsonFile);

    if (!records) {
        return;
    }

    if (!(records instanceof Array)) {

        alert(
            "JSON root must be an array of objects."
        );

        return;
    }


    if (records.length === 0) {

        alert(
            "JSON contains no records."
        );

        return;
    }


    for (var recordIndex = 0; recordIndex < records.length; recordIndex++) {
        if (!isRecord(records[recordIndex])) {
            alert(
                "JSON record " +
                (recordIndex + 1) +
                " must be an object."
            );
            return;
        }
    }


    var templateFile = File.openDialog(
        "Select Illustrator template",
        "*.ai"
    );

    if (!templateFile) {
        return;
    }


    var defaultOutputName = getDefaultOutputName(jsonFile);
    var outputName = prompt(
        "Enter output file name (without extension):",
        defaultOutputName
    );

    if (!outputName) {
        return;
    }

    outputName = outputName.replace(/\.ai$/i, "");
    outputName = safeName(outputName);

    if (outputName === "") {
        alert("A valid output file name is required.");
        return;
    }


    var recordsPerFile = 100;
    var batchCount = Math.ceil(records.length / recordsPerFile);
    var progressWindow = createProgressWindow(records.length);


    for (var batchIndex = 0; batchIndex < batchCount; batchIndex++) {

        var batchStart = batchIndex * recordsPerFile;
        var batchEnd = Math.min(
            batchStart + recordsPerFile,
            records.length
        );
        var doc = app.open(templateFile);


    var placeholders =
        detectPlaceholders(doc);


    if (placeholders.length === 0) {

        alert(
            "No placeholders found.\n\n" +
            "Use format such as:\n" +
            "{{name}}\n" +
            "{{id}}\n" +
            "{{designation}}"
        );

        doc.close(
            SaveOptions.DONOTSAVECHANGES
        );

        return;
    }


    var fieldMessage =
        "Detected placeholders:\n\n";

    for (
        var p = 0;
        p < placeholders.length;
        p++
    ) {

        fieldMessage +=
            "• " +
            placeholders[p] +
            "\n";
    }


    alert(fieldMessage);


    var templateArtboard =
        doc.artboards[0];

    var rect =
        templateArtboard.artboardRect;


    var templateLeft = rect[0];
    var templateTop = rect[1];
    var templateRight = rect[2];
    var templateBottom = rect[3];


    var artboardWidth =
        templateRight -
        templateLeft;


    var artboardHeight =
        templateTop -
        templateBottom;

    var layoutColumns = artboardWidth > artboardHeight
        ? 1
        : 10;


    templateLeft = 0;
    templateTop = artboardHeight;
    templateArtboard.artboardRect = [
        templateLeft,
        templateTop,
        artboardWidth,
        0
    ];


    var sourceItems = [];


    for (
        var i = 0;
        i < doc.pageItems.length;
        i++
    ) {

        var item =
            doc.pageItems[i];


        if (
            item.parent &&
            item.parent.typename === "Layer"
        ) {

            sourceItems.push(item);
        }
    }


    if (sourceItems.length === 0) {

        alert(
            "No artwork found in template."
        );

        doc.close(
            SaveOptions.DONOTSAVECHANGES
        );

        return;
    }


    var outputFile = new File(
        jsonFile.parent.fsName + "/" +
        getOutputName(outputName, batchIndex, batchCount)
    );


    doc.saveAs(
        outputFile,
        new IllustratorSaveOptions()
    );


    for (
        var index = batchStart;
        index < batchEnd;
        index++
    ) {

        var record =
            records[index];


        var column =
            (index - batchStart) % layoutColumns;


        var row =
            Math.floor(
                (index - batchStart) / layoutColumns
            );


        var left =
            templateLeft +
            column *
            (
                artboardWidth +
                CONFIG.horizontalGap
            );


        var top =
            templateTop -
            row *
            (
                artboardHeight +
                CONFIG.verticalGap
            );


        var right =
            left +
            artboardWidth;


        var bottom =
            top -
            artboardHeight;


        var artboard;

        if (index === batchStart) {
            artboard = doc.artboards[0];
        } else {
            artboard = doc.artboards.add([
                left,
                top,
                right,
                bottom
            ]);
        }


        artboard.name =
            createArtboardName(
                record,
                index,
                placeholders
            );


        var targetLayer = null;


        if (CONFIG.separateLayers) {

            targetLayer =
                doc.layers.add();


            targetLayer.name =
                "Item " +
                pad(index + 1, 3);
        }


        var copies = [];


        for (
            var j = 0;
            j < sourceItems.length;
            j++
        ) {

            try {

                var copy =
                    sourceItems[j].duplicate(
                        targetLayer ||
                        doc.activeLayer,
                        ElementPlacement.PLACEATEND
                    );


                copies.push(copy);

            } catch (e) {

            }
        }


        if (copies.length > 0) {

            var bounds =
                getCombinedBounds(
                    copies
                );


            var artworkCenterX =
                (bounds[0] + bounds[2]) / 2;


            var artworkCenterY =
                (bounds[1] + bounds[3]) / 2;


            var targetCenterX =
                (left + right) / 2;


            var targetCenterY =
                (top + bottom) / 2;


            var moveX =
                targetCenterX -
                artworkCenterX;


            var moveY =
                targetCenterY -
                artworkCenterY;


            for (
                var m = 0;
                m < copies.length;
                m++
            ) {

                try {

                    copies[m].translate(
                        moveX,
                        moveY
                    );

                } catch (e) {}
            }
        }


        for (
            var c = 0;
            c < copies.length;
            c++
        ) {

            replaceDynamic(
                copies[c],
                record
            );
        }


        updateProgress(
            progressWindow,
            index + 1,
            records.length,
            batchIndex + 1,
            batchCount
        );


    }


    for (
        var d = sourceItems.length - 1;
        d >= 0;
        d--
    ) {

        try {
            sourceItems[d].remove();
        } catch (e) {}
    }


    doc.save();


    doc.close(SaveOptions.DONOTSAVECHANGES);
    }


    closeProgressWindow(progressWindow);
    alert(
        "DONE!\n\n" +
        "Records: " +
        records.length +
        "\nFiles: " +
        batchCount
    );
}


/**
 * Creates and initializes a floating ScriptUI progress bar palette.
 * @param {number} totalRecords - Total count of records in dataset.
 * @returns {Window} ScriptUI Window instance.
 */
function createProgressWindow(totalRecords) {

    var window = new Window(
        "palette",
        "Generating artwork"
    );

    window.orientation = "column";
    window.alignChildren = "fill";
    window.margins = 16;

    window.status = window.add(
        "statictext",
        undefined,
        "Preparing 0 of " + totalRecords + " records..."
    );

    window.progress = window.add(
        "progressbar",
        undefined,
        0,
        totalRecords
    );
    window.progress.preferredSize = [360, 20];
    window.show();
    window.update();

    return window;
}


/**
 * Updates progress bar position and status messaging with redraw.
 * @param {Window} window     - Active ScriptUI progress window.
 * @param {number} completed  - Processed records count.
 * @param {number} total      - Total records count.
 * @param {number} batch      - Current batch index (1-based).
 * @param {number} batchCount - Total batch files count.
 */
function updateProgress(
    window,
    completed,
    total,
    batch,
    batchCount
) {

    if (!window) {
        return;
    }

    window.progress.value = completed;
    window.status.text =
        "Processing " +
        completed +
        " of " +
        total +
        " records  |  File " +
        batch +
        " of " +
        batchCount;
    window.update();
    app.redraw();
    $.sleep(10);
}


/**
 * Closes and disposes the ScriptUI progress window safely.
 * @param {Window} window - ScriptUI Window instance to close.
 */
function closeProgressWindow(window) {

    if (window) {
        window.close();
    }
}


/**
 * Scans all text frames in the document to detect all unique placeholder keys.
 * @param {Document} doc - Target Illustrator document.
 * @returns {string[]} Array of unique detected placeholder strings.
 */
function detectPlaceholders(doc) {

    var found = [];


    for (
        var i = 0;
        i < doc.textFrames.length;
        i++
    ) {

        var text =
            doc.textFrames[i].contents;


        var matches =
            extractPlaceholders(text);


        for (
            var j = 0;
            j < matches.length;
            j++
        ) {

            var field =
                matches[j];


            if (!containsValue(found, field)) {

                found.push(field);
            }
        }
    }


    return found;
}


/**
 * Parses a string to extract all substring tokens bounded by CONFIG delimiters.
 * Trims whitespace and deduplicates extracted keys.
 * @param {string} text - Raw text string from a TextFrame.
 * @returns {string[]} Array of unique placeholder keys found.
 */
function extractPlaceholders(text) {

    var result = [];


    var start =
        CONFIG.placeholderStart;


    var end =
        CONFIG.placeholderEnd;


    var position = 0;


    while (true) {

        var startIndex =
            text.indexOf(
                start,
                position
            );


        if (startIndex === -1) {
            break;
        }


        var endIndex =
            text.indexOf(
                end,
                startIndex +
                start.length
            );


        if (endIndex === -1) {
            break;
        }


        var field =
            text.substring(
                startIndex +
                start.length,
                endIndex
            );


        field =
            field.replace(
                /^\s+|\s+$/g,
                ""
            );


        if (
            field.length > 0 &&
            !containsValue(result, field)
        ) {

            result.push(field);
        }


        position =
            endIndex +
            end.length;
    }


    return result;
}


/**
 * Checks if an array contains a specific value (case-sensitive comparison).
 * @param {Array} values         - Array to inspect.
 * @param {*}     requestedValue - Value to search for.
 * @returns {boolean} True if found, false otherwise.
 */
function containsValue(values, requestedValue) {

    for (var index = 0; index < values.length; index++) {
        if (values[index] === requestedValue) {
            return true;
        }
    }

    return false;
}


/**
 * Recursively inspects text frames and grouped items to replace placeholder tags with record data.
 * @param {PageItem} item   - PageItem (TextFrame or GroupItem) to inspect.
 * @param {Object}   record - JSON record dictionary for current item.
 */
function replaceDynamic(
    item,
    record
) {

    if (
        item.typename ===
        "TextFrame"
    ) {

        var text =
            item.contents;


        var fields =
            extractPlaceholders(text);


        for (
            var i = 0;
            i < fields.length;
            i++
        ) {

            var field =
                fields[i];


            var value =
                findJSONValue(
                    record,
                    field
                );


            if (value === null) {

                if (
                    CONFIG.missingFieldBehavior ===
                    "blank"
                ) {

                    value = "";

                } else {

                    value =
                        CONFIG.placeholderStart +
                        field +
                        CONFIG.placeholderEnd;
                }
            }


            text =
                replaceAll(
                    text,
                    CONFIG.placeholderStart +
                    field +
                    CONFIG.placeholderEnd,
                    value
                );
        }


        item.contents =
            text;


        return;
    }


    if (
        item.typename ===
        "GroupItem"
    ) {

        for (
            var j = 0;
            j < item.pageItems.length;
            j++
        ) {

            replaceDynamic(
                item.pageItems[j],
                record
            );
        }


        return;
    }
}


/**
 * Resolves a field value from a record using exact match first, falling back to normalized fuzzy match.
 * @param {Object} object         - JSON record data object.
 * @param {string} requestedField - Field name extracted from placeholder.
 * @returns {string|null} Resolved string value, or null if key does not exist.
 */
function findJSONValue(
    object,
    requestedField
) {

    if (
        object.hasOwnProperty(
            requestedField
        )
    ) {

        return normalizeValue(
            object[requestedField]
        );
    }


    var requested = normalizeFieldName(requestedField);


    for (
        var key in object
    ) {

        if (
            object.hasOwnProperty(key)
        ) {

            if (normalizeFieldName(key) === requested) {

                return normalizeValue(
                    object[key]
                );
            }
        }
    }


    return null;
}


/**
 * Normalizes field names for fuzzy matching by lowercasing and stripping non-alphanumeric characters.
 * @param {*} value - Field name string or key.
 * @returns {string} Normalized alphanumeric string.
 */
function normalizeFieldName(value) {

    return String(value)
        .toLowerCase()
        .replace(/[^a-z0-9]/g, "");
}


/**
 * Validates whether an entity is a non-null, non-array JavaScript object.
 * @param {*} value - Entity to check.
 * @returns {boolean} True if valid record object, false otherwise.
 */
function isRecord(value) {

    return value !== null &&
        typeof value === "object" &&
        !(value instanceof Array);
}


/**
 * Converts null/undefined to empty string, serializes objects, and stringifies values for text frames.
 * @param {*} value - Raw value from JSON record.
 * @returns {string} String representation.
 */
function normalizeValue(value) {

    if (
        value === null ||
        value === undefined
    ) {

        return "";
    }


    if (
        typeof value === "object"
    ) {

        try {

            return stringifyJSON(value);

        } catch (e) {

            return String(value);
        }
    }


    return String(value);
}


/**
 * Replaces all occurrences of a target substring within a string (ECMAScript 3 compatible).
 * @param {string} text        - Source string.
 * @param {string} search      - Substring to replace.
 * @param {string} replacement - Replacement string.
 * @returns {string} Resulting string.
 */
function replaceAll(
    text,
    search,
    replacement
) {

    return text.split(search).join(
        replacement
    );
}


/**
 * Calculates aggregate geometric bounding box [left, top, right, bottom] for an array of items.
 * @param {PageItem[]} items - Array of Illustrator page items.
 * @returns {number[]} Combined bounds array in points.
 */
function getCombinedBounds(
    items
) {

    var left = null;
    var top = null;
    var right = null;
    var bottom = null;


    for (
        var i = 0;
        i < items.length;
        i++
    ) {

        try {

            var b =
                items[i].visibleBounds;


            if (left === null) {

                left = b[0];
                top = b[1];
                right = b[2];
                bottom = b[3];

            } else {

                left =
                    Math.min(
                        left,
                        b[0]
                    );


                top =
                    Math.max(
                        top,
                        b[1]
                    );


                right =
                    Math.max(
                        right,
                        b[2]
                    );


                bottom =
                    Math.min(
                        bottom,
                        b[3]
                    );
            }

        } catch (e) {}
    }


    return [
        left,
        top,
        right,
        bottom
    ];
}


/**
 * Determines a filesystem-safe artboard name from the first matched placeholder or sequential fallback.
 * @param {Object}   record       - Current record object.
 * @param {number}   index        - Zero-based index of item.
 * @param {string[]} placeholders - Array of detected placeholder names.
 * @returns {string} Sanitized artboard name.
 */
function createArtboardName(
    record,
    index,
    placeholders
) {

    var name = "";


    for (var fieldIndex = 0; fieldIndex < placeholders.length; fieldIndex++) {
        var fieldValue = findJSONValue(
            record,
            placeholders[fieldIndex]
        );

        if (fieldValue !== null && fieldValue !== "") {
            name = fieldValue;
            break;
        }
    }


    if (name === "") {

        name =
            "Item " +
            pad(index + 1, 3);
    }


    return safeName(name);
}


/**
 * Sanitizes a string for filesystem safety by replacing invalid characters with underscores.
 * @param {string} name - Input string.
 * @returns {string} Sanitized string.
 */
function safeName(name) {

    return String(name)
        .replace(
            /[\\\/:*?"<>|]/g,
            "_"
        )
        .replace(
            /[\r\n]/g,
            ""
        );
}


/**
 * Left-pads a number with leading zeros up to specified length.
 * @param {number|string} number - Number to pad.
 * @param {number}        length - Desired total length.
 * @returns {string} Padded string (e.g. 1 -> "001").
 */
function pad(number, length) {

    var value =
        String(number);


    while (
        value.length < length
    ) {

        value =
            "0" +
            value;
    }


    return value;
}


/**
 * Constructs default base output filename by appending "-output" to JSON filename.
 * @param {File} jsonFile - Source JSON File object.
 * @returns {string} Proposed base output filename.
 */
function getDefaultOutputName(jsonFile) {

    return jsonFile.name.replace(
        /\.json$/i,
        ""
    ) + "-output";
}


/**
 * Formats batch file name, appending "-part-XX.ai" when job is split into multiple files.
 * @param {string} baseName   - Base output filename.
 * @param {number} batchIndex - Zero-based batch index.
 * @param {number} batchCount - Total batch files count.
 * @returns {string} Final filename with .ai extension.
 */
function getOutputName(baseName, batchIndex, batchCount) {


    var suffix = batchCount > 1
        ? "-part-" + pad(batchIndex + 1, 2)
        : "";

    return safeName(
        baseName +
        suffix +
        ".ai"
    );
}


/**
 * Reads and parses UTF-8 JSON file from disk with error reporting.
 * @param {File} file - Illustrator File object to read.
 * @returns {Array|null} Parsed JSON data or null on failure.
 */
function readJSON(file) {
    file.encoding = "UTF-8";

    if (!file.open("r")) {
        alert("Cannot open JSON file.");
        return null;
    }

    var content = file.read();
    file.close();

    try {
        return parseJSON(content);
    } catch (error) {
        alert("JSON parsing failed.\n\n" + error);
        return null;
    }
}


/**
 * Evaluates JSON string into JavaScript objects in ExtendScript ECMAScript 3.
 * @param {string} text - JSON string.
 * @returns {*} Evaluated data structure.
 */
function parseJSON(text) {
    return eval("(" + text + ")");
}


/**
 * Custom recursive JSON stringifier compatible with ExtendScript.
 * @param {*} value - JavaScript object, array, or primitive.
 * @returns {string} JSON-formatted string.
 */
function stringifyJSON(value) {
    if (value === null) {
        return "null";
    }

    if (typeof value !== "object") {
        return String(value);
    }

    if (value instanceof Array) {
        var arrayParts = [];
        for (var index = 0; index < value.length; index++) {
            arrayParts.push(stringifyJSON(value[index]));
        }
        return "[" + arrayParts.join(",") + "]";
    }

    var objectParts = [];
    for (var key in value) {
        if (value.hasOwnProperty(key)) {
            objectParts.push(
                "\"" +
                String(key).replace(/\\/g, "\\\\").replace(/\"/g, "\\\"") +
                "\":" +
                stringifyJSON(value[key])
            );
        }
    }
    return "{" + objectParts.join(",") + "}";
}


/**
 * Top-Level Execution Guard
 * Executes main() within a global try-catch block to display informative alert dialogs on errors.
 */
try {
    main();
} catch (error) {
    alert(
        "ERROR\n\n" +
        error +
        "\n\nLine: " +
        error.line
    );
}
/**
 * extract_text.jsx
 * InDesign CC 2023+ ExtendScript — template file
 *
 * DO NOT run this file directly. The Python batch runner (batch_extract.py)
 * reads this file, injects the two path variables at the top, and sends the
 * combined script to InDesign via COM (DoScript). This avoids all command-
 * line argument issues with modern InDesign versions.
 *
 * The injected header will look like:
 *   var inddPath = "C:\\path\\to\\file.indd";
 *   var jsonPath = "C:\\path\\to\\output.json";
 */

(function () {

    var inddFile = new File(inddPath);
    if (!inddFile.exists) {
        throw new Error("Source file not found: " + inddPath);
    }

    app.scriptPreferences.enableRedraw = false;
    app.scriptPreferences.userInteractionLevel = UserInteractionLevels.NEVER_INTERACT;

    var doc;
    try {
        doc = app.open(inddFile, false); // false = do not show window
    } catch (e) {
        throw new Error("Could not open: " + inddPath + " | " + e.message);
    }

    try {
        var result = extractDocument(doc, inddPath);
        writeJSON(jsonPath, result);
    } finally {
        doc.close(SaveOptions.NO);
        app.scriptPreferences.enableRedraw = true;
        app.scriptPreferences.userInteractionLevel = UserInteractionLevels.INTERACT_WITH_ALL;
    }

})();


// ---------------------------------------------------------------------------
// Extract all stories from the document
// ---------------------------------------------------------------------------
function extractDocument(doc, sourcePath) {

    var meta = {};
    try {
        var xmp      = doc.metadataPreferences;
        meta.title       = xmp.title        || "";
        meta.author      = xmp.author       || "";
        meta.description = xmp.description  || "";
        meta.keywords    = xmp.keywords     || "";
    } catch (e) {
        meta.error = "XMP read error: " + e.message;
    }

    var stories = [];
    for (var i = 0; i < doc.stories.length; i++) {
        var sd = extractStory(doc.stories[i], i);
        if (sd.text.replace(/\s+/g, "") !== "") {
            stories.push(sd);
        }
    }

    return {
        source_file : sourcePath,
        document    : doc.name,
        page_count  : doc.pages.length,
        metadata    : meta,
        story_count : stories.length,
        stories     : stories
    };
}


// ---------------------------------------------------------------------------
// Extract a single story
// ---------------------------------------------------------------------------
function extractStory(story, idx) {

    var pages   = [];
    var pageSet = {};
    for (var f = 0; f < story.textContainers.length; f++) {
        try {
            var pn = story.textContainers[f].parentPage
                   ? story.textContainers[f].parentPage.name
                   : "pasteboard";
            if (!pageSet[pn]) { pageSet[pn] = true; pages.push(pn); }
        } catch (e) {}
    }

    var paragraphs = [];
    for (var p = 0; p < story.paragraphs.length; p++) {
        var para = story.paragraphs[p];
        var txt  = normaliseText(para.contents);
        if (txt.replace(/\s+/g, "") === "") { continue; }
        var style = "";
        try { style = para.appliedParagraphStyle.name; } catch (e) {}
        paragraphs.push({ style: style, text: txt });
    }

    var full = "";
    for (var q = 0; q < paragraphs.length; q++) {
        full += paragraphs[q].text + "\n";
    }

    return {
        story_index     : idx,
        frame_count     : story.textContainers.length,
        pages           : pages,
        paragraph_count : paragraphs.length,
        paragraphs      : paragraphs,
        text            : full.replace(/\n$/, "")
    };
}


// ---------------------------------------------------------------------------
// Normalise InDesign special characters to plain Unicode
// ---------------------------------------------------------------------------
function normaliseText(s) {
    s = s.replace(/\r/g,     "\n");   // paragraph return
    s = s.replace(/\u0003/g, "\n");   // forced line break
    s = s.replace(/\u0004/g, "\t");   // tab
    s = s.replace(/\u2028/g, "\n");   // line separator
    s = s.replace(/\uFFFD/g, "");     // missing glyph
    s = s.replace(/[ \t]{2,}/g, " "); // collapse multiple spaces
    return s;
}


// ---------------------------------------------------------------------------
// Write JSON  (ExtendScript has no native JSON.stringify)
// ---------------------------------------------------------------------------
function writeJSON(jsonPath, obj) {
    var f = new File(jsonPath);
    f.encoding = "UTF-8";
    f.lineFeed = "Unix";
    f.open("w");
    f.write(toJSON(obj, 0));
    f.close();
}

function toJSON(val, depth) {
    var ind  = repeat("  ", depth);
    var ind1 = repeat("  ", depth + 1);
    var t = typeof val;
    if (val === null || val === undefined) { return "null"; }
    if (t === "boolean") { return val ? "true" : "false"; }
    if (t === "number")  { return isFinite(val) ? String(val) : "null"; }
    if (t === "string")  { return jsonString(val); }
    if (val instanceof Array) {
        if (!val.length) { return "[]"; }
        var ap = [];
        for (var i = 0; i < val.length; i++) {
            ap.push(ind1 + toJSON(val[i], depth + 1));
        }
        return "[\n" + ap.join(",\n") + "\n" + ind + "]";
    }
    var keys = [], k;
    for (k in val) { if (val.hasOwnProperty(k)) { keys.push(k); } }
    if (!keys.length) { return "{}"; }
    var op = [];
    for (var j = 0; j < keys.length; j++) {
        op.push(ind1 + jsonString(keys[j]) + ": " + toJSON(val[keys[j]], depth + 1));
    }
    return "{\n" + op.join(",\n") + "\n" + ind + "}";
}

function jsonString(s) {
    s = String(s)
        .replace(/\\/g,  "\\\\")
        .replace(/"/g,   "\\\"")
        .replace(/\n/g,  "\\n")
        .replace(/\r/g,  "\\r")
        .replace(/\t/g,  "\\t")
        .replace(/[\x00-\x1F]/g, function(c) {
            return "\\u00" + ("00" + c.charCodeAt(0).toString(16)).slice(-2);
        });
    return '"' + s + '"';
}

function repeat(s, n) {
    var o = "";
    for (var i = 0; i < n; i++) { o += s; }
    return o;
}

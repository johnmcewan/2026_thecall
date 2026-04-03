/**
 * extract_text.jsx
 * Optimized for Legacy InDesign Files (2008-2014)
 * Handles "Invalid Control Character" errors for large batch processing.
 */

var debugLog = new File(Folder.temp + "/indd_debug.txt");

(function() {
    try {
        // Variables inddPath and jsonPath are injected by the Python runner
        var inddFile = new File(inddPath);
        
        if (!inddFile.exists) {
            throw new Error("Source file not found at: " + inddPath);
        }

        app.scriptPreferences.enableRedraw = false;
        app.scriptPreferences.userInteractionLevel = UserInteractionLevels.NEVER_INTERACT;

        var doc;
        try {
            // OpenOptions.DEFAULT_VALUE is critical for bypassing 2008-2014 conversion dialogs
            doc = app.open(inddFile, false, OpenOptions.DEFAULT_VALUE); 
        } catch (e) {
            throw new Error("InDesign Open Error: " + e.message);
        }

        try {
            var result = extractDocument(doc, inddPath);
            writeJSON(jsonPath, result);
        } catch (e) {
            throw new Error("Extraction/Write Error: " + e.message);
        } finally {
            if (doc) {
                doc.close(SaveOptions.NO);
            }
            app.scriptPreferences.enableRedraw = true;
            app.scriptPreferences.userInteractionLevel = UserInteractionLevels.INTERACT_WITH_ALL;
        }

    } catch (err) {
        debugLog.open("a");
        debugLog.writeln(new Date().toLocaleString() + " | Error in " + inddPath + ": " + err.message);
        debugLog.close();
    }
})();

// ---------------------------------------------------------------------------
// Extraction Logic
// ---------------------------------------------------------------------------

function extractDocument(doc, sourcePath) {
    var meta = {};
    try {
        var xmp = doc.metadataPreferences;
        meta.title = xmp.title || "";
        meta.author = xmp.author || "";
    } catch (e) { meta.error = "Metadata inaccessible"; }

    var pagesResults = [];
    
    for (var i = 0; i < doc.pages.length; i++) {
        var page = doc.pages[i];
        var pageData = { page_name: page.name, frames: [] };

        var textFrames = page.textFrames.everyItem().getElements();
        
        // Sort frames Top-to-Bottom, then Left-to-Right
        textFrames.sort(function(a, b) {
            var b1 = a.geometricBounds; 
            var b2 = b.geometricBounds;
            return (Math.abs(b1[0] - b2[0]) > 3) ? (b1[0] - b2[0]) : (b1[1] - b2[1]);
        });

        for (var j = 0; j < textFrames.length; j++) {
            var tf = textFrames[j];
            var frameContent = extractFrame(tf);
            if (frameContent.text.replace(/\s+/g, "") !== "") {
                pageData.frames.push(frameContent);
            }
        }
        pagesResults.push(pageData);
    }

    return {
        source_file: sourcePath,
        document_name: doc.name,
        page_count: doc.pages.length,
        metadata: meta,
        pages: pagesResults
    };
}

function extractFrame(tf) {
    var paragraphs = [];
    for (var p = 0; p < tf.paragraphs.length; p++) {
        var para = tf.paragraphs[p];
        var txt = sanitizeText(para.contents.toString());
        
        if (txt.replace(/\s+/g, "") === "") continue;
        
        paragraphs.push({
            style: para.appliedParagraphStyle.name,
            text: txt
        });
    }
    return {
        bounds: tf.geometricBounds,
        paragraphs: paragraphs,
        text: sanitizeText(tf.contents.toString())
    };
}

/**
 * Aggressively cleans text of InDesign control characters that break JSON.
 */
function sanitizeText(str) {
    if (!str) return "";
    var s = str.toString();
    // Replace InDesign specific breaks with standard characters
    s = s.replace(/\u0003/g, "\n");   // Forced Line Break
    s = s.replace(/\u000D/g, "\n");   // Carriage Return
    s = s.replace(/\u0007/g, "");     // Indent to Here marker
    s = s.replace(/\u0009/g, " ");    // Tab
    
    // Strip all other non-printable control characters (Unicode 0000-001F)
    s = s.replace(/[\u0000-\u001F\u007F-\u009F\uFEFF\u2028\u2029]/g, "");
    return s;
}

// ---------------------------------------------------------------------------
// JSON Utility (Safe Encoder)
// ---------------------------------------------------------------------------

function writeJSON(path, obj) {
    var f = new File(path);
    f.encoding = "UTF-8";
    f.open("w");
    f.write(toJSON(obj));
    f.close();
}

function toJSON(obj) {
    var parts = [];
    if (obj === null) return "null";
    if (typeof obj === "string") return jsonEscape(obj);
    if (typeof obj === "number" || typeof obj === "boolean") return obj.toString();
    if (obj instanceof Array) {
        for (var i = 0; i < obj.length; i++) parts.push(toJSON(obj[i]));
        return "[" + parts.join(",") + "]";
    }
    for (var key in obj) {
        if (obj.hasOwnProperty(key)) {
            parts.push(jsonEscape(key) + ":" + toJSON(obj[key]));
        }
    }
    return "{" + parts.join(",") + "}";
}

function jsonEscape(s) {
    return '"' + s.replace(/\\/g, "\\\\")
                  .replace(/"/g, '\\"')
                  .replace(/\n/g, "\\n")
                  .replace(/\r/g, "\\r")
                  .replace(/\t/g, "\\t")
                  .replace(/[\u0000-\u001f]/g, function(c) {
                      return "\\u" + ("0000" + c.charCodeAt(0).toString(16)).slice(-4);
                  }) + '"';
}
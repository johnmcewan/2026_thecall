"""
test_indesign2.py  —  probe scriptPreferences and open() behaviour
"""
import pythoncom
import win32com.client
import time
import os

pythoncom.CoInitialize()
app = win32com.client.Dispatch("InDesign.Application.2026")
time.sleep(3)

os.makedirs("C:/temp", exist_ok=True)

# Put the path to ONE of your .indd files here for the open test
TEST_INDD = r"E:\Callproject\assembled\153_Untitled CD\July 27-Aug 02\Pages\Classified Pages_\Classified.indd"

script = r"""
var report = [];

// 1. What properties does scriptPreferences expose?
var sp = app.scriptPreferences;
var spProps = [
    "enableRedraw", "userInteractionLevel", "version",
    "scriptsPanel", "measurementUnit", "numberingList"
];
for (var i = 0; i < spProps.length; i++) {
    try {
        report.push("scriptPreferences." + spProps[i] + ": " + sp[spProps[i]]);
    } catch(e) {
        report.push("scriptPreferences." + spProps[i] + ": MISSING");
    }
}

// 2. Try opening with 3 args: file, showWindow, showOptions
//    showOptions=false is supposed to suppress dialogs
var testFile = new File("__TESTINDD__");
var doc;
var opened = false;
try {
    doc = app.open(testFile, false, false);
    opened = true;
    report.push("open(file, false, false): SUCCESS — " + doc.name);
} catch(e) {
    report.push("open(file, false, false): FAILED — " + e.message);
}

// 3. If that failed, try 2 args
if (!opened) {
    try {
        doc = app.open(testFile, false);
        opened = true;
        report.push("open(file, false): SUCCESS — " + doc.name);
    } catch(e) {
        report.push("open(file, false): FAILED — " + e.message);
    }
}

// 4. If that failed, try 1 arg
if (!opened) {
    try {
        doc = app.open(testFile);
        opened = true;
        report.push("open(file): SUCCESS — " + doc.name);
    } catch(e) {
        report.push("open(file): FAILED — " + e.message);
    }
}

if (opened && doc) {
    doc.close(SaveOptions.NO);
    report.push("close: OK");
}

var f = new File("C:/temp/indesign_probe2.txt");
f.encoding = "UTF-8";
f.open("w");
f.write(report.join("\n"));
f.close();
"done";
""".replace("__TESTINDD__", TEST_INDD.replace("\\", "\\\\"))

try:
    result = app.DoScript(script, 1246973031)
    print("DoScript returned:", result)
    with open("C:/temp/indesign_probe2.txt", "r", encoding="utf-8", errors="replace") as f:
        print(f.read())
except Exception as e:
    print("Error:", e)

pythoncom.CoUninitialize()
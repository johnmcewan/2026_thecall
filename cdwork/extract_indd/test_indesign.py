"""
test_indesign.py  —  run this once to diagnose the InDesign 2026 COM interface
"""
import pythoncom
import win32com.client
import time

pythoncom.CoInitialize()

app = win32com.client.Dispatch("InDesign.Application.2026")
time.sleep(3)

# Send a simple introspection script
script = """
var report = [];
report.push("InDesign version: " + app.version);

// Test which properties exist
var props = [
    "userInteractionLevel",
    "openOptions", 
    "scriptPreferences",
    "activeDocument"
];
for (var i = 0; i < props.length; i++) {
    try {
        var val = app[props[i]];
        report.push(props[i] + ": EXISTS (" + typeof val + ")");
    } catch(e) {
        report.push(props[i] + ": MISSING - " + e.message);
    }
}

// Test open() signature — how many args does it accept?
report.push("app.open type: " + typeof app.open);

// Write results to a temp file
var f = new File("C:/temp/indesign_probe.txt");
f.encoding = "UTF-8";
f.open("w");
f.write(report.join("\\n"));
f.close();
"done";
"""

import os
os.makedirs("C:/temp", exist_ok=True)

try:
    result = app.DoScript(script, 1246973031)
    print("DoScript returned:", result)
    with open("C:/temp/indesign_probe.txt", "r") as f:
        print(f.read())
except Exception as e:
    print("Error:", e)

pythoncom.CoUninitialize()


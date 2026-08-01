TOOLPY IMAGE CAPTURE FEATURE

COPY THESE FILES INTO YOUR TOOLPY FOLDER:
- word_page.py
- clipboard_capture.py

Replace the existing word_page.py.
clipboard_capture.py is a new file.

No extra pip packages are required.

TEST:
1. Run:
   py main.py

2. Select an open Word document in ToolPy.

3. Click Start Capture.

4. Copy images one at a time from:
   - Browsers
   - Paint
   - Screenshots
   - Discord or similar apps
   - File Explorer

5. Watch the captured count increase.

6. Click Paste Images.

ToolPy pastes all captured images at the end of the selected Word document,
clears the temporary captured images, stops recording, and returns to Idle.

REBUILD EXE:
rmdir /s /q build
rmdir /s /q dist
py -m PyInstaller --clean ToolPy.spec

PyInstaller should automatically detect clipboard_capture.py because
word_page.py imports it.

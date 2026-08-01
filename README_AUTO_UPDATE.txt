TOOLPY AUTO-UPDATE SETUP

COPY THESE INTO YOUR TOOLPY FOLDER
- main.py
- updater.py
- ToolPy.spec
- version.txt

Keep your existing:
- assets folder
- theme.py
- utils.py
- word_page.py
- word_tools.py

FIRST BUILD
1. Publish GitHub release v0.1.0 with an asset named exactly:
   ToolPy.exe

2. Build:
   rmdir /s /q build
   rmdir /s /q dist
   py -m PyInstaller --clean ToolPy.spec

HOW TO PUBLISH THE NEXT UPDATE
1. Change version.txt, for example:
   0.1.1

2. Build again:
   rmdir /s /q build
   rmdir /s /q dist
   py -m PyInstaller --clean ToolPy.spec

3. Create a new GitHub release:
   Tag: v0.1.1
   Asset filename: ToolPy.exe

4. Open the older ToolPy.exe on the other PC.
   It will ask to update, download the new EXE,
   replace itself, and restart automatically.

IMPORTANT
- The repository/release must be public.
- The release asset must always be named ToolPy.exe.
- Running `py main.py` skips automatic EXE replacement.

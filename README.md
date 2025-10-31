# coderdoc
Blood Test Parser: Data extraction from PDF or Image (via OCR) to Excel/CSV Spreadsheet with Graph Generation for Analysis

In order to get it working, set your exams files (pdf or images) inside the /exames/ folder. It can work for any language, just translate the terms inside the file compilar_exames.py file.

---- HOW TO RUN THE CODE DIRECTLY (WITH NO EXE COMPILATION) IN LINUX ----

Be sure you are in the coderdoc folder in the prompt, then run:

./compilar_exames.py

---- HOW TO BUILD THE EXE FOR LINUX ----

Be sure you are in the coderdoc folder in the prompt, then run:

python3 -m PyInstaller coderdoc.spec

---- HOW TO BUILD THE EXE FOR WINDOWS ----

1. Install OpenCV:

```powershell
"C:\Users\YOUR_USERNAME\AppData\Local\Programs\Python\Python313\python.exe" -m pip install opencv-python
```

If the script is currently using some extras image funcions, also install this:

```powershell
"C:\Users\YOUR_USERNAME\AppData\Local\Programs\Python\Python313\python.exe" -m pip install opencv-python-headless
```

2. Test the `cv2`:

```powershell
"C:\Users\YOUR_USERNAME\AppData\Local\Programs\Python\Python313\python.exe" -c "import cv2; print(cv2.__version__)"
```

That should print the OpenCV version (ex: `4.10.0`).


3. Just then run the EXE builder to compyle the code for Windows:

```powershell
rmdir /s /q build dist __pycache__
"C:\Users\YOUR_USERNAME\AppData\Local\Programs\Python\Python313\python.exe" -m PyInstaller coderdoc.spec
```

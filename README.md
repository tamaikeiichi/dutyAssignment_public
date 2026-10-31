# Decision algorithm for work schedule using OR-Tools
## files
### dutyAssign.py
main code
### input_1.xx.xlsx
sample input
### input_1.xx_assigned_scorexx.xlsx
sample output
## Get started
To install dependencies of python, run the following commands:

```bash
$ python.exe ./install_packages.py
```

To install Electron and the JavaScript libraries (japanese-holidays, exceljs, tabulator-tables) listed in package.json:

```bash
$ npm.cmd install
```

To build, 

```bash
$ npm.cmd run build:py
$ npm.cmd run build
```

The installer is created as `dist/当直表作成 Setup <version>.exe`, and an unpacked version that runs without installing is created in `dist/win-unpacked/`.

### Adding files loaded at runtime

Only the files listed in `build.files` of package.json are packaged into the app.
If you add a file that the app loads at runtime (a new .js / .html / .css file, an image, or a library file under `node_modules`), add it to `build.files` as well.
Otherwise the packaged app fails to load it, while `npm.cmd start` keeps working because it reads files directly from the project folder.

```jsonc
"files": [
  "main.js",
  "preload.js",
  "renderer.js",
  "index.html",
  "result.html",
  "assets/**/*",
  "!node_modules/**/*",
  "node_modules/japanese-holidays/**/*",
  "node_modules/exceljs/dist/exceljs.min.js",
  "node_modules/tabulator-tables/dist/js/tabulator.min.js",
  "node_modules/tabulator-tables/dist/css/tabulator.min.css",
  "new-file.js"                                  // <- add new files here
]
```

- `node_modules` is excluded as a whole (`!node_modules/**/*`) and only the files actually loaded are listed after it, so add a new library's file (e.g. its `dist/*.min.js`) explicitly.
- A library required from `main.js` (main process) needs its whole folder, like `node_modules/japanese-holidays/**/*`.
- `dutyAssign.exe` is not listed here; it is copied by `extraResources`.

To check after building:

```bash
# list the files packaged into the app
$ npx.cmd @electron/asar list "dist/win-unpacked/resources/app.asar"

# run the packaged app without installing and try the feature that uses the new file
$ "dist/win-unpacked/当直表作成.exe"
```

If the feature does not work only in the packaged app, compare the asar list with the files referenced in the HTML/JS (`<script src>`, `<link href>`, `require()`) and add the missing ones to `build.files`.

Required libraries:

```bash
$ pip list
Package                   Version
------------------------- -----------
absl-py                   2.3.1
altgraph                  0.17.4
et_xmlfile                2.0.0
immutabledict             4.2.2
importlib_metadata        8.7.0
numpy                     2.0.2
openpyxl                  3.1.5
ortools                   9.0.9048  # highly recommended to use this version!
packaging                 25.0
pandas                    2.3.3
pefile                    2023.2.7
pip                       25.2
protobuf                  3.20.3    # highly recommended to use this version!
pyinstaller               6.16.0
pyinstaller-hooks-contrib 2025.9
python-dateutil           2.9.0.post0
pytz                      2025.2
pywin32-ctypes            0.2.3
setuptools                57.4.0
six                       1.17.0
typing_extensions         4.15.0
tzdata                    2025.2
zipp                      3.23.0

$ python --version
Python 3.9.7

$ node --version
v22.21.0

$ npm list
excel-paste-app@1.0.0
├── cross-env@7.0.3
├── dotenv@17.2.3
├── electron-builder@26.0.12
├── electron@28.3.3
└── npm-run-all@4.1.5

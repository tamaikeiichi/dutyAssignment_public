# Decision algorithm for work schedule using OR-Tools
## files
### dutyAssign.py
main code
### input.xlsx
sample input
### assigned_schedule_score12056.0.xlsx
sample output

To install dependencies of python, run the following commands:

```bash
$ python.exe ./install_packages.py
```

To install dependencies of Node.js, run the following commands:

```bash
$ npm install electron
```

Required libraries are descrbed below:

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

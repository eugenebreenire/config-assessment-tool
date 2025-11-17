# CompareResults

The configuration assessment tool (CAT) see [here](https://github.com/Appdynamics/config-assessment-tool) provides metrics on how well your applications are instrumented based on field best practices.

CompareResults piggybacks on the CAT output to compare a previous result against a current result APM, BRUM and MRUM — workbooks ending with “-MaturityAssessment-apm.xlsx”, “-MaturityAssessment-brum.xlsx” or “-MaturityAssessment-mrum.xlsx”.


## Requirements

- Two outputs from the Configuration Assessment Tool
- Python 3.8 or later.
- Internet access to load Chart.js from a CDN.

# Setup!

## Setup Instructions

1. Unzip the `CONFIG-ASSESSMENT-TOOL` folder.

2. On mac - open a terminal and navigate to the `compare-plugin` directory:
    cd path/to/unzipped/CONFIG-ASSESSMENT-TOOL/compare-plugin

3. Run the setup script using bash:
    ./setup.sh

4. After the bash script has complete and if all modules have been installed - run this command:
    source venv/bin/activate

5. Then run this:
    python3 core.py

6. The UI should automatically launch with an address of: http://127.0.0.1:5000/ - see considerations for upload. 
    - The CAT reports we can compare at this time are:
        - APM output - ending with "-MaturityAssessment-apm.xlsx" 
        - BRUM output - ending with "-MaturityAssessment-brum.xlsx
        - MRUM output - ending with "-MaturityAssessment-mrum.xlsx
    - The previous and current reports have to be from the same controller - otherwise the script will terminate
    - For best results ensure the previous report is dated before the Current report
    - When we successfully compare two reports the UI will display 3 outputs:
        - Excel Workbook - this is a low level view of the comparison
        - PPT - this is a high level summary of the comparison
        - JSON - this is used to easily look at the comparison results of a single application in the insights function


## If bash will not run:

- chmod +x setup.sh

## Module Not Found Errors
Modules should be installed as part of setup.sh, however, if you get Module Not Found Errors when running core.py (Error: ModuleNotFoundError: No module named 'openpyxl') you will have to install each Module.

Below is a list of the modules needed:
- Flask
- pandas
- openpyxl
- python-pptx
- xlwings

Below is the versions being used
- Flask>=2.3.2
- pandas>=1.5.3
- openpyxl>=3.1.2
- python-pptx>=0.6.21
- xlwings>=0.28.0

Install Module as follows:
- pip3 install <<module>> or pip install <<module>>
- Example if you receive: Error: ModuleNotFoundError: No module named 'openpyxl' - enter pip3 install openpyxl 


To help with a successful comparison, see below: 
1. APM, BRUM and MRUM workbooks "{jobName}-MaturityAssessment-apm.xlsx" can be compared for now
2. The xlsx files to be compared have to be from the same controller 
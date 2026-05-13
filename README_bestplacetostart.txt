SPINCYCLE INSTRUCTIONS_________________________________https://github.com/stefanhaugen


0. copying the entire folder to Desktop has been my preferred way to run this.
1. This program can be run offline (no Internet, local to instrument computer or your laptop), or online.
2. double click the .bat file that says "SETUP_once.bat". This will run until completion.
3. SETUP just performed installs python (coding language, patch specific) to your computer, and all associated python packages and utilities.
4. anytime you'd like to launch the Data Pipeline interface, double click the .bat file called "launch140SpinCycle.bat" and your default internet browser will open with a local session.


ABOUT PYTHON ON YOUR COMPUTER

This program brings its own python. It puts it in a "python" folder right inside this folder, not on the rest of your computer.

  - if you don't have python installed at all: doesn't matter, this works
  - if you have python installed already: also doesn't matter, this program ignores it
  - this program will NOT overwrite, upgrade, or touch any python you already have
  - if you uninstall this program, just delete the folder. nothing else changes on your computer.

The version this program uses is python 3.11. Specific patch. Picked because it works.


WHAT EACH FILE DOES

  README_bestplacetostart.txt   - this file you're reading right now
  SETUP_once.bat                - run this ONE time, on a computer with internet. installs python + packages into the "python" folder inside this folder
  launch_140SpinCycle.bat       - run this every time you want to use the program. opens it in your browser
  bootstrap_pip.py              - helper for SETUP. SETUP runs it for you. you never touch this
  requirements.txt              - list of python packages SETUP installs. you don't open this
  app__4_.py                    - the program itself. the .bat files run this for you. you don't open this either
  HPLC_Main_Page.py             - placeholder, not used yet
  Multiple_File_Processing.py   - placeholder, not used yet
  Statistical_Analyses.py       - placeholder, not used yet

  packages/                     - folder of installer files SETUP uses. don't delete, don't modify
  python/                       - this folder gets created the first time you run SETUP. it's the embedded python copy. don't open, don't modify


TROUBLESHOOTING
  "Embedded Python not found"
    - Run SETUP_once.bat on an internet-connected machine first

  "Package installation failed"
    - Delete the "python" folder and re-run SETUP_once.bat

  App launches but shows import errors
    - Delete the "python" folder and re-run SETUP_once.bat

  Want a completely fresh start?
    - Delete the "python" folder (everything else stays intact)
    - Re-run SETUP_once.bat


ADVANCED USERS

1. copying the entire parent v1 folder over to an instrument computer (after running SETUP on an internet-connected computer) will run the program locally on the machine without internet access.
Confirmation of this is that they python folder will be populated upon SETUP running. The program has everything it needs after SETUP_once.bat is ran

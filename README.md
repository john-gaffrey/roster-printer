# roster-printer
A helper utility to fetch, filter, format and print roster spreadsheets for a mutli-classroom environment 

Tested on Python 3.11

Usage:
* download source
* copy or rename `config_example.yaml` into `config.yaml` and modify to match your environment
* On Windows, ensure that a program is installed that adds the `print` option to the right-click menu for `.pdf` files.
    * Edge does not work
    * Adobe and Sumatra do.
    * ensure that your desired program is set as the default open for `.pdf`.
* ensure the correct printer is configured as the default in the system settings.
* install python 3 and pip, with or without a venv
    * if using a venv, mnake sure to edit `roster-printer.bat` or `roster-printer.sh`to point to the venv'd `python`
* run `python -m pip install requirements.txt`
* create a shortcut or symlink to `roster-printer.bat` or `roster-printer.sh` in a convenient location for the operators
    * Windows steps:
        * create shortcut
        * edit shortcut target to be: `cmd.exe /c start "" "C:\path\to\roster-printer.bat"`
    * linux steps: 
        * coming soon
* run

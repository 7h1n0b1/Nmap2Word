# Nmap2Word
Convert Nmap's XML output to MS Word table format.

# Usage
Simply download the python file and run it with python3.
> git clone https://github.com/7h1n0b1/Nmap2Word.git

> pip install -r requirements.txt

> python3 nxml.py

*Or you can simply run the nxml.exe*

There are 2 ways in which you may proceed. 
- Initiate a Nmap Scan
- Use XML output report of a previous scan

You should have nmap installed on your local machine. After executing the script choose option 1 to run the scan. Here you can either: 
1. Enter one or multiple targets manually
![Manually Enter Targets](https://github.com/7h1n0b1/Nmap2Word/blob/master/POC/ManualTarg.gif)

2. Select a file having multiple targets (The targets should be newline or/and comma seperated)
![Select File With Multiple Targets](https://github.com/7h1n0b1/Nmap2Word/blob/master/POC/FileTargs.gif)

If you already have an XML output of a scan you ran previously, you can simply select option 2.
Here you can give your XML output to the script and the script will generate the word output report.
![Report from XML File](https://github.com/7h1n0b1/Nmap2Word/blob/master/POC/XMLTarg.gif)

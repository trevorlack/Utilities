# Utilities
Python Utilities to Automate Various Processes


FTP Script for GlobalScape FTP
Coded using python COM library (win32).  To download the target file, the remote folder is first filtered down and then the list of remaining files names are exported to a txt file on the C drive with the CuteFTP "GetList" method.  The resulting text file then goes into a pandas dataframe and is evaluated to find the target file dates.  This additional processes eliminates the need to pull down all files that meet a common criteria then loop through them once they are saved in your local folder.

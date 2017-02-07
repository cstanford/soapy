# soapy
A simple tool for creating SLP SOAP notes in excel. 

### Requirements
* [python3](https://www.python.org/downloads/)
* [openpyxl](https://www.google.com/search?q=openpyxl&oq=openpyxl&aqs=chrome..69i57j0j69i60l3j69i59.2895j0j4&sourceid=chrome&ie=UTF-8)

### How to

Simply run the script passing a filename as a parameter:
```
$ python3 soap.py myfile
```
> Note: Be sure the file you wish to edit is not open in Excel.


Voila! Follow the instructions on screen to create your SOAP template. 

#### Notes:
* To add another patient to the same worksheet, simply run the program again with the same filename. 
* Soapy will overwrite a patient's worksheet if the patient already exists in the workbook. 


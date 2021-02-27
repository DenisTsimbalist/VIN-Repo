VIN-Repo is a simple webscraping tool with two scripts: random_vin.py and vin_getter.py. 

The first sccript, random_vin.py, will generate a list of randomized VINs and their corresponding year, make and model which is returned to an Excel sheet random_vins.xlsx.

The vin_getter.py script is used to get the vehicle's year, make and model from just a VIN. This is returned to vins.xlsx.

Both scripts utilize chromedriver.exe through Selenium. Pyperclip is used in the vin_getter.py script which allows the user to copy a list of VINs to their clipboard (works best if the VINs are copied from an Excel book) and then run their script. 

The random_vin.py script is used as a test for vin_getter.py to confirm that the data being pulled is correct. The thinking being if that the randomized VINs return the correct year, make and model there will be confidence in getting correct data from a list of VINs with unknown year, make and model.

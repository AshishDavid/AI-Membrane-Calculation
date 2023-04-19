# AI-Membrane-Calculation

<a rel="license" href="http://creativecommons.org/licenses/by-sa/4.0/"><img alt="Creative Commons License" style="border-width:0" src="https://i.creativecommons.org/l/by-sa/4.0/88x31.png" /></a><br />This work is licensed under a <a rel="license" href="http://creativecommons.org/licenses/by-sa/4.0/">Creative Commons Attribution-ShareAlike 4.0 International License</a>.

All materials are available under the Creative Commons license "Attribution-ShareAlike" 4.0.
When borrowing any materials from this repository, you must leave a link to it, and also specify my name: Ashish David.


# How to run the program
Install the dependencies. Use the command "pip install -r requirements.txt" on terminal to install it.

Use any Python IDE or terminal.

Use this command to run the program:

python main.py Location json/csv --rso1=rs-0 --amembrane1=amembrane --gas1=H2/CH4/He/CO2/NO2

You can input maximum 5 files at the same time.
For Example: ```python main.py --filename1="D:\Diffusivity calculation\Raw data\H2 raw data.xlsx" --filename2="D:\Diffusivity calculation\Raw data\CO2 raw data.xlsx" csv --rso1=126.9 --rso2=1 --amembrane1=5.73 --amembrane2=5.73 --gas1=H2 --gas2=CO2```

And you will see the result on the screen and also the output will be stored in output_in_json.json file or output_in_csv.csv in the same directory as the program.

### USE INSTRUCTIONS ###

Hello! I know because this is a console application, it looks daunting. But have no fear!

Here's a list of instructions you need to use this program.


#### SETUP ####
First, go to the AZ Software Center and download Python 3.9.

Then, go to the file explore. Right click on the folder containing all of the program's python source files (azemail.py, azsort.py, etc.).
From the menu that pops up, choose 'Open Terminal.'

When the console opens up, make sure you are working in the folder.

Next, copy paste the following into the terminal and hit enter:
python -m pip install --trusted-host pypi.org --trusted-host pypi.python.org --trusted-host files.pythonhosted.org -r requirements.txt

It should be downloading all the necessary packages needed for the program! 

Finally, ensure that you have 
	- a local folder somewhere with all of the emails (.msg files) you want to analyze
	- a spreadsheet where you want to add email logs to

#### RUNNING THE PROGRAM ####

Open the console. Make sure you are working in the same folder as azemail.py.

Copy paste the following into the console and hit enter:
	python azemail.py
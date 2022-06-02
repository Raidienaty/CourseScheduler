# Course Schedule Builder

## Purpose
This project's goal was to build a tool that can take a schedule used by my school's administration and build out individual major schedules. This tool will separate out by major from the excel sheet given and build individual major schedules.

## How to use
If you're on Windows, you're in luck! We have an exe file just for you. If not, I hope you enjoy using a terminal with a python interpretter.

1. If you are familiar with GitHub, please skip this step and just clone the repo. Otherwise, download the project by clicking the green `Code` button above. Once clicked, it will display a menu. Click `Download Zip`. Unpack the zip file into a location you won't lose, consider your desktop.
2. Open up the folder, confirm you have a `Schedules` directory and a `Course Grid UPDATED.xlsx` file along with the `main.exe` file.
3. Place your course report excel document in this folder with the `Course Grid UPDATED.xlsx` and `main.exe` files. 
4. Once you have the report in the folder, double click the `main.exe` file. This will run it.
5. To see the output, please check the `Schedules` folder. This is where all the schedules that are built end up.
6. Once done, please make sure to delete your previous report! If you don't, the program will crash!

## CLI Users

1. Please confirm you have `git` installed on your machine through your respective package manager (apt, dnf, yum, brew, ~~snap~~)
2. Please run the following to clone the repo to your machine:
```
git clone https://github.com/Raidienaty/CourseScheduler.git
```
3. Once downloaded, move your report file into the folder.
4. Run the program using
```sh
python3 main.py
```
5. Be sure to delete your report before running it with a new report. You can only have the one report in the folder at a time.
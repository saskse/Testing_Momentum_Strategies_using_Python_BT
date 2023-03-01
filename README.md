# Testing Momentum Strategies using Python

This is the repository for the Bachelor Thesis of Saskia Senn conducted at UZH Zurich.

## Abstract
The aim of this thesis is to develop an automated correction tool using the Python programming language to efficiently correct the *Involving Activity 3* in the course *Asset Management: Investments*. The exercise requires students to create two momentum strategies, a long-only and a long-short, based on historical stock prices of 18 stocks using varying look-back and holding periods. The tool is designed to be highly flexible in terms of input data, lock-back, and holding periods, enabling the momentum strategies to be effectively tested and compared to a buy-and-hold strategy. The tool offers a powerful approach for correcting the *Involving Activity 3* leading to faster processing times and minimized errors compared to manual correction methods.

## Table of Contents
- [Getting Started](#getting-started)
- [Directory Structure](#directory-structure)
- [Running the Framework](#running-the-framework)
- [Deliverables](#deliverables)

## Getting Started 

### Requirements
- Python 3.10.5 or newer 
- Visual Studio Code 1.74.3 or newer
- GitHub Pull Requests and Issues
- Office 2021 or newer

### Setting up the Repository 
Open a new Visual Studio Code prompt window on Windows (on Linux and MacOS a normal shell will do). Set-up the project repository by doing the following:

1. Sign in with GitHub in the Visual Studio Code Application
2. Use the Clone Repository button in the Source Control view (available when you have no folder open).
3. Insert Repository-URL: https://github.com/saskse/Testing_Momentum_Strategies_using_Python_BT.git
4. Choose a folder as Repository-target

## Directory Structure
- `code`: directory containing source code of the correction
    - `student/`: source files for the student correction
        - `dir_stud.py`
        - `correction_stud.py`
    - `wb/`: source files for the executive education participants correction
        - `dir_wb.py`
        - `correction_wb.py`
- `data/`: main data directory
    - `input/`: folder that holds:
         - IA_Output_empty_stud.xlsx
         - IA_Output_empty_wb.xlsx
         - IA_3_HS22_shifted.xlsx (To calculate the monthly return consistently the data needs to be shifted by one month, starting one week earlier ([see Table 2](deliverables/Bachelor_Thesis_Saskia_Senn (42).pdf.pdf#page=23))
    - `output/`: folder where the IA Output will get exported to
    - `files_for_correction_stud/`: folder with all the submitted files
- `correction_manual`: a manual for the *Headcoach*, who is responsible for correcting the *Involving Activity*, is provided

## Running the Framework
The correction framework allows for two different execution modes: dir_stud, the user will run the correction for the *Involving Acitvity* of the student. In dir_wb, the correction will run for the executive education participants.

### Student correction
The student correction can be started by performing the following three steps:

1. Download the folder with the submitted files from OLAT and save it on the computer.
2. Adjust the path in both, the dir_stud and correction_stud file.
    - dir_stud: path to where the submitted files from OLAT are saved
    - correction_stud: path to where the input data is saved
4. Press "Run the Python-file" button in the dir_stud file. Execution may take a few hours.

### Executive education participants correction
The executive education participants correction can be started by performing the following three steps:

1. Download the folder with the submitted files from OLAT and save it on the computer.
2. Adjust the path in both, the dir_wb and correction_wb file.
    - dir_wb: path to where the submitted files from OLAT are saved
    - correction_wb: path to where the input data is saved
4. Press "Run the Python-file" button in the dir_wb file. Execution may take a few hours.

## Deliverables
The written thesis can be found [here](deliverables/Bachelor_Thesis_Saskia_Senn (42).pdf.pdf) while a correction manual can be found [here](correction_manual.md).

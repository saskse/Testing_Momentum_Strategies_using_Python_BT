# Correction Manual

## How to
The requirements and how to run the framework is described [here](README.md).

## Yearly Updates
The following adjustments are necessery every year, regardless if there are further changes in the *Involving Activity*:
1.	Update return data, shares and dates
    - Discuss the time period and number of shares with the lecturer
        - When a bigger or smaller time period is used, or more or less shares are considered, it may be necessary to make adjustments when saving the following variables:

          ```bash
            number_of_firms = 18
            time_period = 250
          ```
         - Also while loading the workbook range into a dataframe the range string (first argument) needs to be adjusted:
           ```bash
            df_grunddaten_stud = load_workbook_range("C10:U261", ws_grunddaten_stud, with_index=True, index_name="Datum ")
           ```
            Note: This adjustment must be completed for every worksheet.
    - Pull data from datastream
         - An instruction for loading data from datastream is provided in the "TC Team > Knowhow" folder
3.	Update the assignment file and the excel file that needs to be solved

## Changes in the assignment
If there is a change in the points awarded per exercise or the riskfree rate changes, these variables can be easily adjusted in the code to reflect the new values:
```bash
    riskfree = 0

    points_1_1 = 6
    points_2_1 = 6
    points_2_2 = 6
    points_3_1 = 6
    points_3_2 = 6
    points_4 = 6
    points_5 = 15
```
In the event of any modifications to an exercise, it is necessary to update the code to align with the revised requirements.

## Changes in the sheet layout

### Sheet 1: ”Einleitung”
In the event of a change in the layout of the ”Einleitung” sheet, it is important to note that the location of data within the sheet may be affected. As a result, it may be necessary to adjust the `row=` and `column=` arguments while saving the following variables:
```bash
    first_name = ws_einleitung.cell(row=11, column=5).value
    last_name = ws_einleitung.cell(row=13, column=5).value
    matriculation_number = ws_einleitung.cell(row=15, column=5).value
```
This will allow the code to access and manipulate the relevant data in the updated sheet layout.

### Sheet 2: ”Eingabe der Daten”
Similarly, in the event of any changes made to the layout of the ”Eingabe der Daten” sheet, it may be necessary to adjust the `row=` and `column=` arguments while saving the following variables:
```bash
    lookback_period_month = ws_eingabe_der_daten.cell(row=11, column=12).value
    holding_period_month = ws_eingabe_der_daten.cell(row=13, column=12).value
    mittel_ranking = ws_eingabe_der_daten.cell(row=22, column=11).value
    aktien_lo = ws_eingabe_der_daten.cell(row=26, column=12).value
    aktien_ls = ws_eingabe_der_daten.cell(row=28, column=12).value
```

### *IA Output*
For generating the IA Output and in the case that there has been an update to the layout of the IA Output Excel file, the corresponding arguments `row=` and `column=` in the code must be modified to match the new layout:
```bash
ws_IA_output.cell(row=6+stud_number, column=3).value = matriculation_number
    ws_IA_output.cell(row=6+stud_number, column=4).value = first_name
    ws_IA_output.cell(row=6+stud_number, column=5).value = last_name
    ws_IA_output.cell(row=6+stud_number, column=6).value = points_stud
    ws_IA_output.cell(row=6+stud_number, column=7).value = olat_name
    if points_stud >= max_points*0.4: #passing limit
        ws_IA_output.cell(row=6+stud_number, column=8).value = '1' #means passed
    else: ws_IA_output.cell(row=6+stud_number, column=8).value = '0' #means failed
    ws_IA_output.cell(row=6+stud_number, column=10).value = points_stud_1_1
    ws_IA_output.cell(row=6+stud_number, column=11).value = points_stud_2_1
    ws_IA_output.cell(row=6+stud_number, column=12).value = points_stud_2_2
    ws_IA_output.cell(row=6+stud_number, column=13).value = points_stud_3_1
    ws_IA_output.cell(row=6+stud_number, column=14).value = points_stud_3_2
    ws_IA_output.cell(row=6+stud_number, column=15).value = points_stud_4
    ws_IA_output.cell(row=6+stud_number, column=16).value = points_stud_5
```

The passing limit, which is currently set to 40%, can be adjusted by increasing or decreasing this number as desired.
```bash
 if points_stud >= max_points*0.4: #passing limit
```

## Mass Evaluation on OLAT
After running the code to generate the IA Output, the Headcoach can open the file and copy the columns ”OLAT” and ”Bestanden.” These columns can then be used to
perform a mass evaluation of the students’ results on OLAT.

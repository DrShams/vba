# How to use conditions,formulas and VBA macros in Excel 2016
## I Using conditions
We have 2 files: 
- [x] 1st of them the customers report which represent the book with reports for each day which
- [x] 2nd file represent parameters of mud
1) Let us **open** the first **.xls file**
2) **Select** the range which will be under our conditions. In our case it will be the **Range M66:Y150**
3) **Home -> Conditional Formatting -> New rule**
In the dialogue message pick 
- [x] **Format only cells that contain**
- [x] **Format only cells with -> Cell Value -> equal to -> =0**
![Screen #1](https://github.com/DrShams/vba/blob/main/Step1_paint_empty_cells.png)
## II Using fomulas
![Screen #2](https://github.com/DrShams/vba/blob/main/Step2_sync%20formulas.png)

## III Using the macro

#What is the function of this macro?
1) This macro offers a choice of one of 3 actions, when you run it, we get the following questions in the form of dialog boxes:
	1. Do you want to check the parameters of the drilling fluid in the current report? (submacro mud_checker)
	2. Do you want to create a new day? (submacro new_day)
	3. Do you want to check all parameters of the drilling fluid and their compliance with the programm in all reports? (submacro mud_checker + function for checking all pages)
#2) How is the functionality implemented?
+ Before us is a range of cells in which the design and actual parameters of the drilling fluid are located:
-density
-viscosity
-water efficiency
-MBT
-Gels 10 sec 10 min and so on.

+ The macro does the following:
+ Within programm values:
-divides the values ​​of fields containing signs greater than or equal to, greater than, less, ranges (for example, the range of conditional viscosity from and to), and also divides the Gels values ​​into component parts
- removes extra spaces, filters cells with text, skips cells with a simple dash "-", as well as empty cells
-define the minimum and maximum values ​​of acceptable parameters by writing this data to a two-dimensional array
+ Within the actual values:
-compares the actual parameters of the drilling fluid with the mud programm
- paints cells with parameters inappropriate to the mud programm ones in red
* note: in case of only 1 out of 2 values of Gels either 10 sec or 10 minutes colors the cell orange
+ Separately creates a new day
* to do this, you must be on the active tab of the last day from which the formulas will be copied and replaced
-when creating a new day, copies the formulas of the actual values ​​for the next day and breaks through the actual values ​​for the current day, this is done
to prevent data loss in the event of a change in the parameters of the drilling fluid from a separate excel file "drilling fluid parameters.xlsx" dependent on the daily report)

# How to run this macro?
1) You need to go to the macro editor (keyboard shortcut Alt + F8)
2) Enter any letter and click on [Create] (this is how we create our new macro)
3) Delete the original content, copy and paste the code with the macro
4) Add Regular Expression Library: [Tools -> References -> Microsoft VBScript Regular Expressions 5.5 -> ok]
-This must be done, without this action the macro will not work
5) Before running the macro:
+ In case of creating a new day of the report, be sure to go to the tab on the last day from which the values ​​and formulas will be copied
+ If you need to check the parameters of a particular day, go to the corresponding tab
6) Run the macro (keyboard shortcut Alt + F8) -> [Run]
7) Select the required actions

# What is the main task of this macro?
-Optimization of time for correct filling and checking of daily reports (for mud engineers and their supervisors)

# Can this macro be used for other forms of daily reports?
-Yes, for this you need to edit the source code of the macro and the lines:
    range_plan = "AC63: AC150" 'here you need to select a range of cells that contain design parameters
    range_fact = "M63: Y150" 'here you need to select the range of cells in which our actual parameters will be located

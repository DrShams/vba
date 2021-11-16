#В чем заключается работа данного макроса?
1) Данный макрос предлагает выбор одного из 3 действий, при его запуске мы получаем в виде диалоговых окон следущие вопросы:
	1. Хотите ли вы проверить параметры бурового раствора в текущем рапорте? (подмакрос mud_checker)
	2. Хотите ли вы создать новый день? (подмакрос new_day)
	3. Хотите ли вы проверить все параметры бурового раствора и их соответвствие с проектными во всех рапортах? (подмакрос mud_checker + функция проверки всех страниц)
2) Как реализован функционал?
	+Перед нами диапазон ячеек в которых располагаются проектные и фактические параметры бурового раствора:
		-плотность
		-вязкость
		-водоотдача
		-MBT
		-СНСы и тд.
	+Макрос выполняет следующие действия:
		+В пределах проектных значений:
			-делит значения полей содержащих знаки больше либо равно, больше, меньше, диапазоны (к примеру диапазон условной вязкости от и до), а также делит значения СНС на составные части
			-делает очистку от лишних пробелов, фильтрует ячейки с текстом, пропускает ячейки с простым символом тире "-", а также пустые ячейки
			-выявляет минимальные и максимальные значения допустимых параметров записывая эти данные в двумерный массив
		+В пределах фактических значений:
			-сравнивает фактические параметры бурового раствора с проектными
			-окрашивает ячейки с параметрами несоответствующими проектными красным цветом*
			*примечание: в случае вылета только 1 из 2 значения СНС окрашивает ячейку в оранжевый цвет
		+Отдельно создает новый день 
		*для этого необходимо находиться на активной вкладке последнего дня с которого будет осуществлено копирование и замена формул
		-при создании нового дня копирует формулы фактических значений на следующий день и пробивает фактические значения на текущий день, это делается 
		для предотвращения потерь данных в случае изменения параметров бурового раствора с зависимого от суточного рапорта отдельного excel файла "параметры бурового раствора")
	
#Как запустить данный макрос?
1) Необходимо перейти в редактор макросов (сочетание клавиш Alt+F8)
2) Вписать любую букву и нажать на [Create] (таким образом мы создаем наш новый макрос)
3) Удалить исходное содержимое, копировать и вставить код с макросом
4) Добавить библиотеку регулярных выражений: [Tools -> References -> Microsoft VBScript Regular Expressions 5.5 -> ok]
	-Это нужно сделать обязательно, без данного действия макрос работать не будет
5) Перед запуском макроса:
	+В случае создания нового дня рапорта обязательно перейти по вкладке на последний день с которого будет осуществлено копирование значений и формул
	+Если необходимо проверить параметры какого-то определенного дня перейти на соответствующую вкладку
6) Запустить макрос (сочетание клавиш Alt+F8) -> [Run]
7) Выбрать необходимые действия

#Какую основную задачу выполняет данный макрос?
-Оптимизация времени для корректного заполнения и проверки суточных рапортов

#Можно ли использовать данный макрос для других форм суточных рапортов?
-Да, для этого необходимо отредактировать исходный код макроса а имменно строки:
    range_plan = "AC63:AC150" 'здесь необходимо выбрать диапазон ячеек в которых содержатся проектные параметры
    range_fact = "M63:Y150" ' здесь нужно выбрать диапазон ячеек в которых будут располагаться наши фактические параметры
	
#What is the function of this macro?
1)This macro offers a choice of one of 3 actions, when you run it, we get the following questions in the form of dialog boxes:
	1. Do you want to check the parameters of the drilling fluid in the current report? (submacro mud_checker)
	2. Do you want to create a new day? (submacro new_day)
	3. Do you want to check all parameters of the drilling fluid and their compliance with the mud programm in all reports? (submacro mud_checker + function for checking all pages)
2) How is the functionality implemented?
	+ Before us is a range of cells in which the design and actual parameters of the drilling fluid are located:
		- density
		- viscosity
		- water efficiency
		- MBT
		- Gels 10 sec 10 min and so on.
	+ The macro does the following:
	+ Within programm values:
		- divides the values ​​of fields containing signs: greater than or equal to, greater than, less, ranges (for example, the range of conditional viscosity from and to), and also divides the Gels values ​​into component parts 10 sec and 10 minutes
		- removes extra spaces, filters cells with text, skips cells with a simple dash sign "-", as well as empty cells
		- define the minimum and maximum values ​​of acceptable parameters by writing this data to a two-dimensional array
	+ Within the actual values:
		- compares the actual parameters of the drilling fluid with the mud programm
		- paints cells with parameters inappropriate to the mud programm ones in red
	* note: in case of only 1 out of 2 values of Gels either 10 sec or 10 minutes colors the cell orange
	+ Separately creates a new day
	* to do this, you must be on the active tab of the last day from which the formulas will be copied and replaced
	- when creating a new day, copies the formulas of the actual values ​​for the next day and breaks through the actual values ​​for the current day, this is done
	to prevent data loss in the event of a change in the parameters of the drilling fluid from a separate excel file "drilling fluid parameters.xlsx" dependent on the daily report)

# How to run this macro?
1) You need to go to the macro editor (keyboard shortcut Alt + F8)
2) Enter any letter and click on [Create] (this is how we create our new macro)
3) Delete the original content, copy and paste the code with the macro
4) Add Regular Expression Library: [Tools -> References -> Microsoft VBScript Regular Expressions 5.5 -> ok]
- This must be done, without this action the macro will not work
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

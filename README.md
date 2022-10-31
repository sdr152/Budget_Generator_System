# Budget_Generator_System
A small-scale system designed for a small company that generates a budget from items stored a database. 
/ Un sistema de baja escala para una mediana empresa que genera presupuestos de materiales guardados en una base de datos.

This system is designed to generate budgets for a small-business dedicated to electrical services in Tegucigalpa, Honduras. 
The budget is generated from a preselected list of materials stored in a local database that comes along with the system. 

Initialially, the database is simply a csv file that will migrate to an sql file in a later version.

This program is created using Python. The front-end is purely made up using the Tkinter Module to create the GUI. For the 
back-end, the Pandas Module is used to generate and manipulate dataframes before updating the database; the csv module is
necessary for saving our data and backup copy. Also, several basic math and cost principles are a implemented for this system 
to fully work.

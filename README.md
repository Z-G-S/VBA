# VBA
Contains a VBA project named "Automated Sales Data Entry and Visualisation System" modelled for a small wallet manufacturing business (case study).

# Background & Rationale

The application is an Automated Sales Data Entry and Visualisation System that increases data entry efficiency and displays various visualisations and statistics to aid decision-making. This application has been built for a small business that manufactures hand-made wallets in a variety of colours and personalisation options. The rationale for creating such an application is that automating data entry and outputting insightful analytics will help improve the efficiency, security and usability of the businesses sales data. This could be realised through cost savings and more effective decision-making by using the analytics and visualisations outputted.

# Images of the Application

![Screenshot 2024-02-10 000726](https://github.com/Z-G-S/VBA-Automated-Sales-Data-Entry-and-Visualisation-System/assets/140622522/229aa51d-f1af-4256-9148-3e03a2a640a3)
![Screenshot 2024-02-10 000802](https://github.com/Z-G-S/VBA-Automated-Sales-Data-Entry-and-Visualisation-System/assets/140622522/894e44d2-3cdf-49d0-89fa-7d80bdb8476d)
![Screenshot 2024-02-10 000813](https://github.com/Z-G-S/VBA-Automated-Sales-Data-Entry-and-Visualisation-System/assets/140622522/f4f46ea9-ebde-4ab9-8317-e622caab6377)
![Screenshot 2024-02-10 000821](https://github.com/Z-G-S/VBA-Automated-Sales-Data-Entry-and-Visualisation-System/assets/140622522/37fcd535-ee92-4803-9635-bcc4fb2044c0)
![Screenshot 2024-02-10 000830](https://github.com/Z-G-S/VBA-Automated-Sales-Data-Entry-and-Visualisation-System/assets/140622522/4aa87e31-ebcc-48f1-82ba-44465dded3d6)
![Screenshot 2024-02-10 000838](https://github.com/Z-G-S/VBA-Automated-Sales-Data-Entry-and-Visualisation-System/assets/140622522/32f6e583-ea43-42bd-b540-6228b9d325d7)
![Screenshot 2024-02-10 000849](https://github.com/Z-G-S/VBA-Automated-Sales-Data-Entry-and-Visualisation-System/assets/140622522/6ae261b1-2296-4abe-a926-615abc426a06)
![Screenshot 2024-02-10 000858](https://github.com/Z-G-S/VBA-Automated-Sales-Data-Entry-and-Visualisation-System/assets/140622522/f4a8ee46-fe79-4272-b53b-4e9878b3ab0d)
![Screenshot 2024-02-10 000907](https://github.com/Z-G-S/VBA-Automated-Sales-Data-Entry-and-Visualisation-System/assets/140622522/a224781e-d380-4064-bd72-c7ad133f9c30)
![Screenshot 2024-02-10 000919](https://github.com/Z-G-S/VBA-Automated-Sales-Data-Entry-and-Visualisation-System/assets/140622522/c24eeb0b-aa13-405e-ab06-3dd7e495296b)

# Microsoft Excel Objects/Worksheets

**ThisWorkbook** -
Contains the subroutine to hide all data sheets upon start up and activates the Login userform. It is the first step of enhancing the security of the application.

**Login** -
The login sheet contains no data or subroutine, with its sole purpose of being a bridge between login and accessing the data sheets, acting as a placeholder. 

**Data_Entry and OpenApplication_Click** -
This sheet contains the command button that when clicked opens the application. This sheet also contains various charts and visualisations that are linked to other sheets, and display in the multiple pages of the ASDEVS userform such as the Visualisations page and the World Sales Map page. It also contains a live picture of data from the calculations sheet and so on its own acts as a visualisation dashboard, but also is a necessary store for the visualisations displayed in the ASDEVS userform. 

**Prices_Costs_Data** -
This sheet contains all of the prices and costs data of the products, with editable and non-editable fields clearly labelled.

**Support_Data** -
This sheet provides support data for the application, providing use to various comboboxes and a store for the last edit functionality. 

**Completed** -
This sheet is a database of the associated data collected from a successful entry.

**Cancelled** -
This sheet is a database of the associated data collected from a cancelled entry.

**Calculations** -
The calculations sheet is where the majority of the calculations needed for output analytics and visualisations in the ASDEVS userform and data entry sheet occur. Any changes here will systematically impact how the application deals with values and how/where it outputs them.

# Userforms

_LoginForm_ - Provides a layer of security to the application, only allowing users with correct login credentials to proceed. 
User Authentication and Workbook/Data Sheets access
cmdSignIn_Click: This subroutine executes upon clicking the Sign-In button. It checks for empty fields in the LoginForm textboxes, validates the user credentials, and if correct, makes the data sheets visible to the user, if not, it displays an error. This subroutine enhances security by showing the data sheets only upon correct entry, protecting the data from unauthorised users.

ADSEVS - Contains the bulk of the operations of the application from manipulating data sheets to displaying visualisations and analytics. 
Initialisation and Setup
userform_initialize: Initialises the user form with necessary configurations and data.
ListBoxPages_Change: Manages the list box selection changes.
SetNormalViewZoom: Sets the zoom level of the active worksheet.

Form Control and Navigation
DataEntry_Click: Shows the Data Entry page.
Visualisations_Click: Handles the display and update of charts and financial metrics.
WorldSalesMap_Click: Displays the world sales map.
UserManual_Click: Displays the user manual.
ResetButtonColours: Resets the background colours of specific buttons on the user form.
CloseButton_Click: Closes the user form.

Data Entry, Editing, and Updates
cmdSubmit_Click: Handles the submission of data in the user form with validations.
cmdEdit_Click: Manages the editing of selected entries.
cmdDelete_Click: Handles the deletion of selected entries.
cmdCancel_Click: Manages the cancellation of selected entries.
ResetButton_Click: Clears all input fields and updates counters.
RowCount_Completed: Calculates and updates the number of populated rows in the 'Completed' sheet.
RowCount_Cancelled: Calculates and updates the number of rows populated in the 'Cancelled' sheet.

Search Functionality
cmdSearchID1_Click: Searches for a customer ID in the 'Completed' list.
cmdNextMatch1_Click: Finds the next match in the 'Completed' list.
cmdSearchID2_Click: Searches for a customer ID in the 'Cancelled' list.
cmdNextMatch2_Click: Finds the next match in the 'Cancelled' list.


Sorting Data
cmdNewtoOld1_Click: Sorts data in the 'Completed' sheet from newest to oldest Date.
cmdOldtoNew1_Click: Sorts data in the 'Completed' sheet from oldest to newest Date.
cmdNewtoOld2_Click: Sorts data in the 'Cancelled' sheet from newest to oldest Date.
cmdOldtoNew2_Click: Sorts data in the 'Cancelled' sheet from oldest to newest Date.

Visualisation and Reporting
RenderCharts: Calls subroutines to update the charts.
UpdateAndActivateChart: Refreshes specified charts.
RevenuePieChart_Click, ProfitPieChart_Click, CostPieChart_Click, SourcePieChart_Click: Handle the display of various charts.
RefreshChartsandData_Click: Updates charts and data.
RefreshWorldMap_Click: Updates the world map chart.

Utility and Validation Functions
ValidateComboBox: Validates ComboBox inputs.
GetPriceAndCost: Retrieves price and cost information for products.

Data Export
cmdExport_Click: Exports data to a xlsm file.
ExportSheet: A helper subroutine for data export.

Miscellaneous
cmdFullScreen_Click: Manages full screen mode for the form.
info1_Click, info2_Click: Display the user manual.

Order_Cancellation - Provides the ability to cancel orders.
Initialisation and Setup
userform_initialize: Initialises the form, populates the DateComboBox, sets up an acceptable range that the “Reason” combobox will accept from the Support_Data sheet, and obtains the current date and user information from the system environment.

Form Control and Navigation
cmdSubmit_Click: Handles the submission of data in the user form with validations, transferring the data to the “Cancelled” data sheet.
CloseButton_Click: Closes the userform.
ResetButton_Click: Resets the input fields excluding customer name and email as they're taken from the “Completed” data sheet.

Miscellaneous
info3_Click: Displays the user manual.

User_Manual - Contains a series of pages & images aimed at helping users.
Initialisation and Setup
userform_initialize: Highlights the DataEntryCB button if the multipage value equals zero.

Form Control and Navigation
ResetButtonColours: Used to reset all button colours back to default.
BackHelpButton_Click: Ensures that the menu buttons on the ASDEVS userform are highlighted correctly, and closes the User_Manual.
Subroutines DataEntryCB_Click to CalculationsCB_Click: Perform the necessary page changes and highlighting changes. 

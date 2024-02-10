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

ThisWorkbook
Contains the subroutine to hide all data sheets upon start up and activates the Login userform. It is the first step of enhancing the security of the application.

*Login*
The login sheet contains no data or subroutine, with its sole purpose of being a bridge between login and accessing the data sheets, acting as a placeholder. 

Data_Entry and OpenApplication_Click
This sheet contains the command button that when clicked opens the application. This sheet also contains various charts and visualisations that are linked to other sheets, and display in the multiple pages of the ASDEVS userform such as the Visualisations page and the World Sales Map page. It also contains a live picture of data from the calculations sheet and so on its own acts as a visualisation dashboard, but also is a necessary store for the visualisations displayed in the ASDEVS userform. 

Prices_Costs_Data
This sheet contains all of the prices and costs data of the products, with editable and non-editable fields clearly labelled.

Support_Data
This sheet provides support data for the application, providing use to various comboboxes and a store for the last edit functionality. 

Completed
This sheet is a database of the associated data collected from a successful entry.

Cancelled
This sheet is a database of the associated data collected from a cancelled entry.

Calculations
The calculations sheet is where the majority of the calculations needed for output analytics and visualisations in the ASDEVS userform and data entry sheet occur. Any changes here will systematically impact how the application deals with values and how/where it outputs them.

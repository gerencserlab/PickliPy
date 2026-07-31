# Blue table picklist generator
## Description
Press the blue Play P (pipeline) toolbar button to run this pipeline. This pipeline runs the Blue table picklist generator in Python.

Picklist generator for acoustic dispensing using Beckman Coulter Echo 650 controlled by Revvity PlateWorks. The Echo picklist defines what volumes to dispense from which source plate barcode / well to which destination plate barcode / well. The goal of this script is to help translating experimental designs as 2-dimensional plate maps into a series of pick commands. The script generates the picklist and accompanying inventory and process files from a plate map and an inventory worksheet in an Excel file. Plate maps allow user control over random-access dispensing one or multiple compounds per destination well, at the desired concentrations. The script supports using multiple source and destination plates, volume checking and distributing larger total volume additions to originate from multiple source wells. Both the destination and source plate will move on the shortest possible path within the instrument. The added volumes are rounded to 2.5nL increments, with warnings if this results in inaccuracy. The current plate maps support 384-well format, our using the top left of the same map for 96-well destination plates.

Citation for PickliPy:
Varunya M. Kattunga, Steven A. Wrobel, Chad A. Lerner, Victor M. Derycz, Elizabeth B. Stephens, Ian S. Brown, Hao Cheng, Sima Taghizadeh, Josef Byrne, Susan Gross, Susan Schneider, Chatura Senadheera, Asia Davis-Castillo, Shane Vistalli-Alvarado, Elena Goncharova, John C. Newman, Brianna J. Stubbs, Simon Melov, Gordon Lithgow, Lisa M. Ellerby, Julie K. Andersen and Akos A. Gerencser. Advanced Open-source Experimental-Design Tools for Microplate-Based Assays with Acoustic Liquid Handling. BIORXIV/2026/735934

## Parameters
| # | Name | Type | Description |
|---|------|------|-------------|
| 0 | Picklist name (will be saved as) | Text | Picklist base name (no extension). Output files use this prefix. |
| 1 | Inventory file name (Excel) | Text | Path to SRC inventory Excel file. |
| 2 | Blue tables folder | Text | Folder containing destination blue-table .xlsx files (and optionally *_platesurvey.xml). |


## Structure
![structure](/img/Blue_table_picklist_generator.jpg)

[Image Analyst MKII](https://www.imageanalyst.net) pipeline - saved by V4.3.7 (build 1054)


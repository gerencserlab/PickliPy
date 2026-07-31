# Visualizer
## Description
Press the blue Play P (pipeline) toolbar button to run this pipeline. This pipeline runs the Picklist Visualizer in Python. Results will be saved in a subfolder where the picklist file is.

Picklist generator for acoustic dispensing using Beckman Coulter Echo 650 controlled by Revvity PlateWorks. The Echo picklist defines what volumes to dispense from which source plate barcode / well to which destination plate barcode / well. The goal of this script is to help translating experimental designs as 2-dimensional plate maps into a series of pick commands. The script generates the picklist and accompanying inventory and process files from a plate map and an inventory worksheet in an Excel file. Plate maps allow user control over random-access dispensing one or multiple compounds per destination well, at the desired concentrations. The script supports using multiple source and destination plates, volume checking and distributing larger total volume additions to originate from multiple source wells. Both the destination and source plate will move on the shortest possible path within the instrument. The added volumes are rounded to 2.5nL increments, with warnings if this results in inaccuracy. The current plate maps support 384-well format, our using the top left of the same map for 96-well destination plates.

Citation for PickliPy:
Varunya M. Kattunga, Steven A. Wrobel, Chad A. Lerner, Victor M. Derycz, Elizabeth B. Stephens, Ian S. Brown, Hao Cheng, Sima Taghizadeh, Josef Byrne, Susan Gross, Susan Schneider, Chatura Senadheera, Asia Davis-Castillo, Shane Vistalli-Alvarado, Elena Goncharova, John C. Newman, Brianna J. Stubbs, Simon Melov, Gordon Lithgow, Lisa M. Ellerby, Julie K. Andersen and Akos A. Gerencser. Advanced Open-source Experimental-Design Tools for Microplate-Based Assays with Acoustic Liquid Handling. BIORXIV/2026/735934

## Parameters
| # | Name | Type | Description |
|---|------|------|-------------|
| 0 | Picklist file name (csv) | Text | Path to the picklist CSV created by picklipy |
| 1 | Source plate format | Text | Use 384 when only top-left wells are used but the source plate is physically 384-well. |
| 2 | Destination plate format | Text | Use 384 when only top-left wells are used but the source plate is physically 384-well. |
| 3 | Image format (png, pdf or svg) | Text | Matplotlib output format, usually png, svg, or pdf. |
| 4 | DPI | Text | Output resolution in dots per inch. |
| 5 | Matplotlib colormap name | Text | Matplotlib colormap name for dispense order. |
| 6 | Color scale scope (global, destination) | Text | Use one color gradient across the full picklist ('global') or restart the gradient per destination plate ('destination') |
| 7 | Plate color (black, white) | Text | Plate background palette (default: black, optional: white) |


## Structure
![structure](/img/Visualizer.jpg)

[Image Analyst MKII](https://www.imageanalyst.net) pipeline - saved by V4.3.7 (build 1054)


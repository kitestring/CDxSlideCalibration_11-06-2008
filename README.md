# CDxSlideCalibration_11-06-2008

Date Written: 11/06/2008

Industry: Medical Device Manufacturer

Device: Blood Analyzer

Market: Veterinary

Platform: Electrolyte Sensor Panel – Na+, K+, Ca2+

GUI:
The “Catalyst Dx_XL7111_Rev C.xls” file contains the GUI.

Sample Raw Data File:
89411_CDX3625_H22_1.csv – 8 Buffers are measured across 3 units.  Each measurement yields a single *.csv file.

Template File:
“Catalyst Dx_XL7114_Rev C.xls”

Sample Output:
The “894110-Calibration.xls” file is an example of the output.

Application Description:
QC technicians measure the concentration of 8 buffers using a given lot of electrolyte sensors.  The measurements are done in triplicate across three units.  For each measurement a *.csv file is exported from the unit and dumped into a predefined network directory.  The GUI file contains inputs for various Meta data as well as the file path to the *.csv files.  This application mines the spectrometry data from 24 csv files and dumps them into the template file.  The template file calculates the calibration constants and generates a pass/fail result based upon predefined criteria determined by R&D.  This allows QC technicians to determine if a given lot of sensors meets criteria which determine if it can be sold.  Prior to this application, these calculations and determinations were manually done by members of the R&D team.

# Excel XPS analysis code "Ctrl+Q"

##Abstract
We have developed the XPS and XAS data analysis code served for the synchrotron radiation soft x-ray beamline users to calibrate energy and intensity in photons and electrons. Users easily identify the atomic elements and its chemical shifts based on the XPS and XAS spectra measured at the synchrotron beamline. The relative sensitivity factors for each element are automatically calibrated with photo ionization cross section at a photon energy used for XPS measurement. The basic curve fit function quickly evaluates the atomic ratios or chemical states of elements among samples analyzed. The code works in the Visual Basic Applications (Personal Workbook Macro) on the Microsoft Excel 2007 or later of the Windows version to utilize the Solver Excel add-in for least square fit optimisation. Users easily handle and share the XPS and XAS data on their laptop computers for efficient usage of beamtime, and analyze it into the publications.

##Background
We have developed the XPS and XAS analytical software based on the Excel VBA code, because no software is available for SR-based XPS in public, even though a number of XPS analytical software are available such as MultiPak (ULVAC PHI), Avantage (Thermo Fisher), Spectral analyser (XPS International) etc. They are basically optimised for standalone XPS with AlKa anode, which generates a photon energy at 1486.6 eV. SR-based XPS needs to calibrate the photon energy and intensity for each data. To handle the SR-based XPS data, the professional analytical software such as Igor Pro, Origin, MatLAB, Mathmatica etc has to be used to normalise the photoelectron energy and intensity at the beamline. However, these software are expensive, so limited in license, and elaborate to learn how to use and streamline the analytical process in their own script codes.

MS Excel is a default standard spreadsheet-based software to inspect the numerical and visualised data in various fields of work everyday, because relational formulation among cells in spreadsheet and Solver non-linear optimisation function in Excel add-in robustly handle scientific and engineering data as well as financial and accounting data. In addition, the visual basic applications (VBA) code streamlines the standard data analysis process without elaborate copy and paste actions on the spreadsheet. Even though the optimisation performance and numerical accuracy are quite limited in the big data analysis, our Excel VBA code makes your data analysis simplified for your preliminary XPS and XAS experience on your own PC at experimental site. Ii is also noted that Excel is popular for data analysis purpose because it is based on graphical user interface. R is a free software included with robust analytical and statistical functions, but it is CUI-based software. To promote the academic and industrial research projects, the data will be shared with users and collaborators during or after the experiment to discuss the data quality and necessity for the further experiment. Post-processed Excel data file is easily distributed without any script code attached, because all relational formulations are kept in the spreadsheet. In this report, we present the detailed function of our Excel VBA code optimised for soft x-ray based XPS and XAS.

##Introduction
Ctrl+Q is a powerful code for XPS data analysis based on the Microsoft Excel VBA and solver function. Ctrl+Q has useful functions for energy and intensity calibration, spectral normalisation, peak identification, spectral comparison, background subtraction, peak fitting, and export the summary of fitting results. The code has been developed at the BL3.2Ua in the Siam Photon Laboratory, and optimised for ascii text data from the LabVIEW-based DAQ software. However, any data formatted in the Excel can be analysed by using the code. Various peak shape can be used in the fitting with a number of background functions. The SR-based XPS can vary the photon energy to increase the spectral intensity in a way that the photoionisation cross section increases as the photon energy decreases. The Excel XPS package includes the core level binding energy, chemical shifts for main peaks, and atomic sensitivity factor for each level based on the XPS standard reference used with AlKa anode. The photon energy dependent atomic sensitivity factor is calculated with photoionisation cross section, which also includes in the package. Auger electron energies are listed from the XPS AlKa spectral database to identify the Auger peaks. These database are used for XAS analysis in the soft x-ray energy range as well.

##Installation
The code is based on the VBA, and installed in the VBE as a Personal Workbook Macro with shortcut key assignment. Solver function also needs to be installed in Excel as a default add-in and registered in the VBE for curve fit procedure. Run the code "CLAM2" from the Macro menu or assign the shortcut key "Ctrl" & "q" to run the "CLAM2" from the Option of Macro menu.

##Data loading
Any data opened in the Excel spreadsheet can be analysed in the code. The energy and intensity data in the spreadsheet are prepared in the two columns started from A2 and B2 cells. “KE/eV” at the A1 cell in the same sheet makes the first column as the kinetic energy scale. "BE/eV" at the A1 cell is recognized as the binding energy to the first column data, "PE/eV" at the A1 cell the photon energy, and so on. The workbook must be saved as a name represent for a spectrum data, and then run the code. The code makes several sheets additional to the original sheet named after the workbook filename such as Graph_filename and Fit_filename.

##Energy and intensity calibrations
Standard sample data is used to calibrate the peak energy or normalisation factor. The code has a function to compare the data processed, so the you can easily identify the calibration factors. The package also provides the standard element binding energy and sensitivity data to calibrate the energy.

##Curve fitting
The peaks identified in the energy calibration are processed in the curve fitting with their sensitivity factors at a photon energy you used. Peak area is calibrated with analytical and numerical ways. The number of peaks can be chosen with parameters such as energy, width, amplitude etc. All the parameters can be constrained or limited in a specific range. Amplitude ratio and peak energy difference are also set up in the cell with specified syntax.

###Type of background subtraction and peak fitting function
- Gaussian, Lorentzian, and its blended function with tail parameters for asymmetry
- Doniac-Sunjic and Ulrik Gelius profiles for asymmetric peak
- Pseudo-Voigt based on either sum or product form between Gauss and Lorentz
- Shirley and Tougaard backgrounds with and without spline numerical convolution
- Arctangent and Victoreen backgrounds for XAS pre-edge subtraction
- Double exponential background
- Peak area with various sensitivity factors including photoionization cross section, source angle correction, MFP, analyzer transmission function etc.
- Fermi edge fitting with the Gaussian-convoluted Fermi-Dirac function
- Trapezoidal integration for peak area
- Interpolation of data
- Normalization of spectrum by a reference spectrum
- Multiple file analysis based on the initial parameters used in a file

##Multiple data file analysis
Once the XPS data is analyzed, you can apply the same analysis conditions in the another files in the energy and intensity calibration or fitting curve as initial parameters. All processed data are summarized in the single Excel sheet to evaluate the atomic element ratio.

##Notes
Ctrl+Q has been used for many users during the experiment and post-data processing to publish the data in the manuscript in peer-reviewd journals in the following.

- http://dx.doi.org/10.1016/j.snb.2013.12.017
- http://dx.doi.org/10.1016/j.apsusc.2016.01.180
- http://dx.doi.org/10.1016/j.carbon.2015.01.018
- http://dx.doi.org/10.1016/j.jenvman.2015.09.036

Currently, the package including database is distributed only for users and workshop participants at SLRI. However, the user-defined database workbook (UD.xlsx) is automatically generated in the directory specified in the code to add your elements and its relative sensitivities for AlKa.

###References for database
Database files are available only in PC at the beamline because the software is only licensed in PC at the beamline.
However, the demo or preview version of software can be downloaded and installed to refer the database.
On-line database links are also freely available for everyone to identify the database references.

- http://www.uksaf.org/data.html

####XPS
X-ray data booklet
- http://xdb.lbl.gov/

Values compiled by Gwyn P. Williams (updated Excel file and poster available)
- https://userweb.jlab.org/~gwyn/

Scofield photoionization cross-section database combined with x-ray booklet binding energy database
"Hartree-Slater subshell photoionization cross-sections at 1254 and 1487 eV"
J. H. Scofield
Journal of Electron Spectroscopy and Related Phenomena, 8129-137 (1976).
- http://dx.doi.org/10.1016/0368-2048(76)80015-1

####AES
"Calculated Auger yields and sensitivity factors for KLL-NOO transitions with 1-10 kV primary beams"
S. Mroczkowski and D. Lichtman, J. Vac. Sci. Technol. A 3, 1860 (1985).
- http://dx.doi.org/10.1116/1.572933
- http://www.materialinterface.com/wp-content/uploads/2014/11/Calculated-AES-yields-Matl-Interface.pdf

(Electron beam energy at 1, 3, 5, and 10 keV for relative cross section and derivative factors)


####WebCross folder
Photoionization cross section online database files should be downloaded and placed in this folder.
- https://vuo.elettra.eu/services/elements/WebElements.html

"Atomic Calculation of Photoionization Cross-Sections and Asymmetry Parameters"
J.J. Yeh, Gordon and Breach Science Publishers, Langhorne, PE (USA), 1993.

"Atomic subshell photoionization cross sections and asymmetry parameters: 1 <= Z <= 103"
J.J. Yeh and I.Lindau, Atomic Data and Nuclear Data Tables, 32, 1-155 (1985).
- http://dx.doi.org/10.1016/0092-640X(85)90016-6

Note that database are supposed to be revised and updated locally based on the experiment.
All the database files are based on AlKa source energy at 1486.6 eV, and webCross data normalize the RSF.
You may also check spectral lines and profiles in the link below;

NIST X-ray Photoelectron Spectroscopy Database
- http://srdata.nist.gov/xps/

"The NIST X-ray photoelectron spectroscopy (XPS) database"
C. D. Wagner, NIST Technical Note 1289 (1991).
- https://archive.org/details/nistxrayphotoele1289wagn

The Surface Analysis Society of Japan: Common Data Processing System
- http://www.sasj.jp/COMPRO/index.html


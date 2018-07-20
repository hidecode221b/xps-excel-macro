# Excel XPS and NEXAFS analysis code "Ctrl+Q"

## Abstract
I have developed the XPS and XAS data analysis code for the synchrotron radiation (SR) soft x-ray beamline users to calibrate the energy and intensity in photons and electrons and analyze the peak composition and shape in the background subtracted profile. The code works in the Visual Basic Applications (VBA) on the Windows Microsoft Excel 2007 or later version utilizing the Solver Excel add-in for least square fit optimisation. Excel 2016 on Mac is also supported partly. Users easily handle and share the XPS and XAS data on their laptop computers for efficient usage of beamtime, and analyze it into the publications. According to the peak energy database specified in the code, users identify the atomic elements and its chemical shifts in the XPS and XAS spectra measured at the synchrotron beamline. The relative sensitivity factor and photo ionization cross section database are used to calibrate the peak intensity measured at the synchrotron-based XPS as well as standalone XPS system. The curve fit function sequentially evaluates the atomic or chemical ratios of the elements in your XPS data among the samples. 

## Background
I have developed the Excel visual basic applications (VBA) code to analyze the soft X-ray photoemission spectroscopy (XPS) and absorption spectroscopy (XAS or NEXAFS) spectra, because no software or code was available for SR-based XPS data analysis in public. Even though a number of commercial software have been developing such as CasaXPS (used in Kratos), MultiPak (ULVAC PHI), Avantage (Thermo VG Fisher), Spectral analyser (XPS International), they are basically optimized for the standalone XPS system utilized with the Mg or Al Ka anode as a X-ray source, which generates a photon energy of 1486.6 eV. However, synchrotron radiation (SR) produces a wide spectrum of photon energy and the monochromator and mirror optics deliver SR to the spot on the sample surface in various energy resolution and spot size used for XPS. To process the number of SR-based spectral data, the professional analytical software such as Igor Pro, Origin, MATLAB, Mathmatica etc. has been used to inspect the photoelectron energy and intensity at the beamline prior to the chemical analysis detailed. However, these software are relatively expensive for XPS beginners and further required for scripting to streamline spectral analyses in their own macro language.

Microsoft Excel is a default standard spreadsheet-based software to inspect and visualise the numerical data in various fields of daily jobs, because relational formulation among cells in spreadsheet and Solver non-linear optimisation function in Excel add-in robustly handle scientific and engineering data as well as financial and accounting data. In addition, the VBA code streamlines the standard data analysis process without elaborated copy and paste actions on the spreadsheet, and instantly plots the charts in the worksheet. Even though the optimisation performance and numerical accuracy are quite limited in the big data analysis, the Excel VBA code makes your spectral data analysis simple and comprehensive on your own laptop PC. Python and R are scientific scripting languages included with robust analytical and statistical libraries, and sophisticated integrated development environment (IDE) for each language is freely available. However, to promote the academic and industrial research projects, the data have to be shared with users and collaborators during or after the experiment to discuss the data quality and necessity for the further additional experiment. Excel workbooks analyzed in this code are easily distributed without any script code attached in the workbook, because all relational formulations are stored in the worksheets. Charts in Excel worksheet can also be placed in the Word document or PowerPoint slide as objects for your brief experimental report. In this report, we present the detailed function of our Excel VBA code optimised for soft x-ray based XPS and XAS.

## Introduction
Ctrl+Q is a powerful code for XPS and XAS data analyses based on the Microsoft Excel VBA and solver function. Ctrl+Q has useful functions for energy and intensity calibration, spectral normalisation, peak identification, spectral comparison, background subtraction, peak fitting, and export the summary of fitting results. The code has been developed and optimized for user service at the BL3.2Ua in the Siam Photon Laboratory. Any spectral data imported and formatted in the Excel spreadsheet, which consist at least of spectral intensity and its corresponded energy in two columns, can be analysed by using the code. Various peak shapes can be decomposed in the fitting with a number of background functions. The SR-based XPS used to vary the photon energy to increase the spectral intensity in a way that the photoionisation cross section increases as the photon energy decreases or increases at the resonance energy. The Excel VBA code works well together with the database of core level binding energy, chemical shifts for main peaks, and atomic sensitivity factor for each level based on the XPS standard reference used with Al Ka anode. The photon energy dependent atomic or relatice sensitivity factor (RSF) is evaluated with photoionisation cross section database, which is available online. Auger electron peaks appear in the XPS spectra at the constant kinetic energy, and their binding energies are varied with the photon energy used to measure XPS. The VBA code distinguishes between the XPS and Auger peaks instantly, so we can easily tune the photon energy appropriate for your measurement without overlap among them. For XAS data analysis, XPS database are used for XAS analysis limited in the soft x-ray energy range.  

## Installation
The code is based on the VBA and installed in the VBE as a Personal Workbook Macro with a shortcut key assignment at "CLAM2" listed in the top of Sub procedures. Solver function also needs to be installed in Excel as a default add-in and referenced in the VBE for curve fit procedure. Any data analysis process starts from the Option of Macro menu "CLAM2" or the shortcut key "Ctrl" & "q" assigned even in the blank workbook. In the workbook, the code analyzes the data depending on either what worksheet is active or what syntax is specified in active worksheet. See YouTube Video: https://youtu.be/tWpcnDjkHzo.

## Data loading
Any data formatted in the Excel spreadsheet can be analysed in the code as follows. The energy and intensity data in the spreadsheet are prepared in the two columns below A2 & B2 cells, respectively. “KE/eV” at the A1 cell in the same sheet makes the first column data as the kinetic energy of spectrum. "BE/eV" at the A1 cell is recognized as the binding energy to the first column data, "PE/eV" at the A1 cell the photon energy, and "ME/eV" for any other purposes. The workbook must be saved as a worksheet name, which represents for a spectrum data, and then run the code in the worksheet called as the Data sheet. The code makes several sheets additional to the Data sheet named after the Data sheet or workbook name such as Graph_*filename* and Fit_*filename*. 

What kind of sheet or data column set generated depends on the operation in the worksheet. In the Graph sheet of XPS data, one spectrum data consists of three columns. First column is used for kinetic energy, second binding energy, and third spectral intensity. In the bottom of first three-column set, the other three columns are followed to be further normalized. First and second columns of second set are the same as those in the first set, and third column the normalized spectral intensity by offset and multiple factors. First chart shown in the Graph sheet is based on the second and third columns in the second set, and second chart is on the first and third columns in the second set.

## Comparing data
You can compare the spectral data with one after another in the Graph sheet. Open the Excel file analyzed in this code and type "comp" in the D1 cells of Graph sheet. Choose the Excel files including the Graph sheet in it to be compared. You can also add the data one after another to type "comp" in the every 3 columns after D1 like G1, J1, and so on. Compared data normalized and calibrated in that sheet is easily exported in the sets of two-column data for each spectrum by "exp" in A1 cell. The resulting data in the Exp sheet can be imported in the software you want to plot for further analysis or quality to journal submission. 

## Energy and intensity calibration
Standard sample data or database is used to calibrate the peak energy or spectral intensity. In the Graph sheet, the photon energy, work function, and charging factor are adjustable parameters for binding energy calibration. The offset and multiple factors are also available to scale spectral intensity for data comparison. To compare the multiple spectra at a glance, both ends of spectral intensity are automatically scaled in the syntax of "auto" at A1 cell in Graph sheet. To specify the energy ranges for spectral offset and multiple scaling, "auto[x0,x1:x2,x3]" can be used in a way that the spectral range between x0 and x1 is averaged to be 0 (offset) and the range between x2 and x3 is averaged to be 1 (multiple). If either x0/x2 or x1/x3 is null, nothing happens to be scaled in the corresponding range. The original data scales for offset and multiple are 0 and 1 respectively.

## Spectral normalization
Spectral intensity is divided (normalized) by the other reference spectrum to compensate the noise or contamination happened during the measurement. Reference data can be added as the second data set by "comp" in the Graph sheet prior to the normalization as mentioned above. The code started with "norm" at A1 cell in the Graph sheet continues normalizing the first data set by the second data set leading to the resultant third data set. According to the normalized data in the third data set in the Graph sheet, the Norm sheet is produced with normalized data set to be analyzed in the code further. 

## Curve fitting
The peaks calibrated and identified by the database in the Graph sheet are analyzed in the Fit sheet based on the least-square regression method. Peak area is evaluated with analytical and numerical ways together with the choice of background subtraction processes. The number of peaks can be chosen with parameters such as curve shape, energy, FWHM width, amplitude etc. All the parameters can be constrained or limited in a specific range. Amplitude ratios and peak energy differences among the peaks are also adjustable in the cells with specified syntax. 

### Type of background subtraction and peak fitting function
- Gaussian, Lorentzian, and its blended function with tail parameters for asymmetric peak ("G" to "TSGL")
- Shirley and Tougaard backgrounds blended with and without polynomial function
- Constant, linear, quadratic, and cubic for polynomial background
- Polynomial and its blended backgrounds optimized with peak fittings (active mode: "BG" to "ABG")
- Arctangent and Victoreen backgrounds for XAS pre-edge subtraction
- Trapezoidal (numerical) integration for peak areas normalized by various sensitivity factors including photoionization cross section, source angle correction, MFP, analyzer transmission function etc.
- User-defined function can be easily implemented in the visual Basic programming.

## Multiple data file analysis
Once a XPS data is analyzed in the Excel VBA code, you can apply its analysis in the another Excel files in terms of the energy and intensity calibration or fitting curve as initial parameters. Atomic concentration in each sample for each element is summarized in the Ana sheet based on the Fit sheets of all spectra. The trend of atomic concentrations in samples is summarized in the Rto sheet based on the Ana sheets for all samples.

## Notes
Ctrl+Q has been used for many users during the experiment and post-data processing to publish the data in the manuscript in peer-reviewed journals in the following.

- http://dx.doi.org/10.1016/j.snb.2013.12.017
- http://dx.doi.org/10.1016/j.apsusc.2016.01.180
- http://dx.doi.org/10.1016/j.carbon.2015.01.018
- http://dx.doi.org/10.1016/j.jenvman.2015.09.036
- http://dx.doi.org/10.1039/C6RA09972F

Details of advanced function available in the code will be described in the future.

### References for database
Database files are available only in PC at the beamline because the software is only licensed in PC at the beamline. You can also purchase the complete packages including databases or advanced functions from the link below.

- https://sites.google.com/site/xpsdataanalysispackage/

On-line database links are also freely available for everyone to identify the database references.

- http://www.uksaf.org/data.html

#### XPS
X-ray data booklet
- http://xdb.lbl.gov/

Values compiled by Gwyn P. Williams (updated Excel file and poster available)
- https://userweb.jlab.org/~gwyn/

Scofield photoionization cross-section database combined with x-ray booklet binding energy database
"Hartree-Slater subshell photoionization cross-sections at 1254 and 1487 eV"
J. H. Scofield, Journal of Electron Spectroscopy and Related Phenomena, 8129-137 (1976).
- http://dx.doi.org/10.1016/0368-2048(76)80015-1

#### AES
"Calculated Auger yields and sensitivity factors for KLL-NOO transitions with 1-10 kV primary beams"
S. Mroczkowski and D. Lichtman, J. Vac. Sci. Technol. A 3, 1860 (1985).
- http://dx.doi.org/10.1116/1.572933
- http://www.materialinterface.com/wp-content/uploads/2014/11/Calculated-AES-yields-Matl-Interface.pdf

(Electron beam energy at 1, 3, 5, and 10 keV for relative cross section and derivative factors)

#### XAS
"X-Ray Interactions: Photoabsorption, Scattering, Transmission, and Reflection at E = 50-30,000 eV, Z = 1-92"
B. L. Henke, E. M. Gullikson, and J. C. Davis, Atomic Data and Nuclear Data Tables 54, 181-342 (1993).
- https://doi.org/10.1006/adnd.1993.1013
- http://henke.lbl.gov/optical_constants/asf.html

#### WebCross folder
Photoionization cross section online database files should be downloaded and placed in this folder.
- https://vuo.elettra.eu/services/elements/WebElements.html

"Atomic Calculation of Photoionization Cross-Sections and Asymmetry Parameters"
J. J. Yeh, Gordon and Breach Science Publishers, Langhorne, PE (USA), 1993.

"Atomic subshell photoionization cross sections and asymmetry parameters: 1 <= Z <= 103"
J. J. Yeh and I. Lindau, Atomic Data and Nuclear Data Tables, 32, 1-155 (1985).
- http://dx.doi.org/10.1016/0092-640X(85)90016-6


Note that database are supposed to be revised and updated locally based on the experiment.
All the database files are based on AlKa source energy at 1486.6 eV, and webCross data normalize the RSF.
You may also check spectral lines and profiles online below;

NIST X-ray Photoelectron Spectroscopy Database
- http://srdata.nist.gov/xps/

"The NIST X-ray photoelectron spectroscopy (XPS) database"
C. D. Wagner, NIST Technical Note 1289 (1991).
- https://archive.org/details/nistxrayphotoele1289wagn

The Surface Analysis Society of Japan: Common Data Processing System
- http://www.sasj.jp/COMPRO/index.html


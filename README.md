# Excel XPS analysis code "Ctrl+Q"

##Abstract
I have developed the XPS and XAS data analysis code for the synchrotron radiation (SR) soft x-ray beamline users to calibrate the energy and intensity in photons and electrons and analyze the peak composition and shape in the background subtracted profile. The code works in the Visual Basic Applications (VBA) on the Windows Microsoft Excel 2007 or later version utilizing the Solver Excel add-in for least square fit optimisation. Users easily handle and share the XPS and XAS data on their laptop computers for efficient usage of beamtime, and analyze it into the publications. According to the peak energy database specified in the code, users identify the atomic elements and its chemical shifts in the XPS and XAS spectra measured at the synchrotron beamline. The relative sensitivity factor and photo ionization cross section database are used to calibrate the peak intensity measured at the synchrotron-based XPS as well as standalone XPS system. The curve fit function sequencially evaluates the atomic or chemical ratios of the elements in your XPS data among the samples. 

##Background
I have developed the Excel visual basic applications (VBA) code to analyze the soft X-ray photoemission spectroscopy (XPS) and absorption spectroscopy (XAS) spectra, because no software or code was available for SR-based XPS data analysis in public. Even though a number of commertial software have been developing such as CasaXPS (used in Kratos), MultiPak (ULVAC PHI), Avantage (Thermo VG Fisher), Spectral analyser (XPS International), they are basically optimized for the standalone XPS system utilized with the Mg or Al Ka anode as a X-ray source, which generates a photon energy of 1486.6 eV. However, synchrotron radiation (SR) produces a wide spectrum of photon energy and the monochromator and mirror optics deliver SR to the spot on the sample surface in various energy resolution and spot size used for XPS. To process the number of SR-based XPS data, the professional analytical software such as Igor Pro, Origin, MatLAB, Mathmatica etc. have been used to normalize the photoelectron energy and intensity at the beamline prior to the detailed analysis due to the intrinsic tunability of photon energy and flux. However, these software are relatively expensive for XPS beginners and further required for scripting to streamline spectral analyses in their own syntax.

Microsoft Excel is a default standard spreadsheet-based software to inspect and visualise the numerical data in various fields of work everyday, because relational formulation among cells in spreadsheet and Solver non-linear optimisation function in Excel add-in robustly handle scientific and engineering data as well as financial and accounting data. In addition, the VBA code streamlines the standard data analysis process without elaborated copy and paste actions on the spreadsheet, and instantly plots the charts in the worksheet. Even though the optimisation performance and numerical accuracy are quite limited in the big data analysis, the Excel VBA code makes your XPS data analysis simple and comprehensive on your own laptop PC. Python and R are scientific scripting languages included with robust analytical and statistical libraries, and graphical user interface platform for each language is freely available. However, to promote the academic and industrial research projects, the data have to be shared with users and collaborators during or after the experiment to discuss the data quality and necessity for the further additional experiment. VBA-processed Excel workbooks are easily distributed without any script code attached in the workbook, because all relational formulations are stored in the worksheets. In this report, we present the detailed function of our Excel VBA code optimised for soft x-ray based XPS and XAS.

##Introduction
Ctrl+Q is a powerful code for XPS and XAS data analyses based on the Microsoft Excel VBA and solver function. Ctrl+Q has useful functions for energy and intensity calibration, spectral normalisation, peak identification, spectral comparison, background subtraction, peak fitting, and export the summary of fitting results. The code has been developed and optimized for user service at the BL3.2Ua in the Siam Photon Laboratory. Any spectral data imported and formatted in the Excel spreadsheet, which consist at least of spectral intensity and its corresponded energy in two columns, can be analysed by using the code. Various peak shapes can be decomposed in the fitting with a number of background functions. The SR-based XPS used to vary the photon energy to increase the spectral intensity in a way that the photoionisation cross section increases as the photon energy decreases or increases at the resonance energy. The Excel VBA code works well together with the database of core level binding energy, chemical shifts for main peaks, and atomic sensitivity factor for each level based on the XPS standard reference used with AlKa anode. The photon energy dependent atomic sensitivity factor is evaluated with photoionisation cross section database, which is available online. Auger electron peaks appear in the XPS spectra at the constant kinetic energy, and their binding energies are varied with the photon energy used to measure XPS. The VBA code distinguishes between the XPS and Auger peaks instantly, so we can easily tune the photon energy appropriate for your measurement without overlap among them. XPS database are also used for XAS analysis in the soft x-ray energy range.

##Installation
The code is based on the VBA, and installed in the VBE as a Personal Workbook Macro with shortcut key assignment. Solver function also needs to be installed in Excel as a default add-in and registered in the VBE for curve fit procedure. Run the code "CLAM2" from the Macro menu or assign the shortcut key "Ctrl" & "q" to run the "CLAM2" from the Option of Macro menu.

##Data loading
Any data formatted in the Excel spreadsheet can be analysed in the code as follows. The energy and intensity data in the spreadsheet are prepared in the two columns started from A2 and B2 cells. “KE/eV” at the A1 cell in the same sheet makes the first column as the kinetic energy scale. "BE/eV" at the A1 cell is recognized as the binding energy to the first column data, "PE/eV" at the A1 cell the photon energy, and "ME/eV" for any other purposes. The workbook must be saved as a name represent for a spectrum data, and then run the code. The code makes several sheets additional to the original sheet named after the workbook filename such as Graph_*filename* and Fit_*filename*.

##Comparing data
You can compare the data with another data both analyzed in the code. Open the Excel file analyzed in this code and type "comp" in the D1 cell, and then run the code. Choose the Excel files to be compared. You can also add the data one after another to type "comp" in the cell like G1, J1, and so on.

##Energy and intensity calibrations
Standard sample data is used to calibrate the peak energy or normalisation factor. The code has a function to compare the data processed, so the you can easily identify the calibration factors. The code also generates the template for standard element binding energy and sensitivity data to calibrate the energy. 

##Normalization
Spectral intensity is normalized with the other reference spectrum. "Norm" in A1 cell of Graph sheet and run the code to choose the Excel file to be used for reference.

##Curve fitting
The peaks identified in the energy calibration are processed in the curve fitting with their sensitivity factors at a photon energy you used. Peak area is calibrated with analytical and numerical ways. The number of peaks can be chosen with parameters such as energy, width, amplitude etc. All the parameters can be constrained or limited in a specific range. Amplitude ratio and peak energy difference are also set up in the cell with specified syntax.

###Type of background subtraction and peak fitting function
- Gaussian, Lorentzian, and its blended function with tail parameters for asymmetry
- Shirley and Tougaard backgrounds with and without spline numerical convolution
- Arctangent and Victoreen backgrounds for XAS pre-edge subtraction
- Peak area with various sensitivity factors including photoionization cross section, source angle correction, MFP, analyzer transmission function etc.
- Fermi edge fitting with the Gaussian-convoluted Fermi-Dirac function
- Trapezoidal integration for peak area
- Multiple file analysis based on the initial parameters used in a file

##Multiple data file analysis
Once the XPS data is analyzed, you can apply the same analysis conditions in the another files in the energy and intensity calibration or fitting curve as initial parameters. All processed data are summarized in the single Excel sheet to evaluate the atomic element ratio.

##Notes
Ctrl+Q has been used for many users during the experiment and post-data processing to publish the data in the manuscript in peer-reviewd journals in the following.

- http://dx.doi.org/10.1016/j.snb.2013.12.017
- http://dx.doi.org/10.1016/j.apsusc.2016.01.180
- http://dx.doi.org/10.1016/j.carbon.2015.01.018
- http://dx.doi.org/10.1016/j.jenvman.2015.09.036

Details of advanced function available in the code will be described in the future.

###References for database
Database files are available only in PC at the beamline because the software is only licensed in PC at the beamline. You can also purchase the complete packages including databases or advanced functions from the link below.

- https://sites.google.com/site/xpsdataanalysispackage/

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


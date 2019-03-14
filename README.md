# Excel XPS and NEXAFS analysis suite "Ctrl+Q"

## Abstract
I have developed the XPS and XAS data analysis suite (so-called "Ctrl+Q") for the synchrotron radiation (SR) soft x-ray beamline users to calibrate the energy and intensity in photons and electrons and analyze the peak composition and shape in the background subtracted profile. The code in the suite works in the Visual Basic for Applications (VBA) on the Windows Microsoft&copy; Excel 2007 or later version utilizing the Solver Add-In for curve-fit optimisation. Excel 2016 on Mac or later version is also partly supported. Users easily handle and share the XPS and XAS data on Excel spread-sheet, and analyze it into the publications with your collaborators. Users immediately identify the atomic elements and its chemical shifts in their measured XPS and XAS spectra during the beamtime without annoying the calibration of energy and intensity at the SR beamline. The relative sensitivity factor and photo ionization cross section are taken into account to evaluate the atomic and chemical ratios. The various fitting functions in both peak and background are available to optimize the variables under certain constraints. User-defined function is also easily implemented by VBA scripting to untangle the spectral complexities in the sample characterization.

## Introduction
Ctrl+Q has been developed on Excel Visual Basic for Applications (VBA) platform to analyze the soft X-ray photoemission spectroscopy (XPS) and absorption spectroscopy (XAS or NEXAFS) spectra, because no open-source program was available for SR-based XPS data analysis. Even though a number of software have been developing in commercial base such as CasaXPS&copy; (used in Kratos), MultiPak (ULVAC PHI), Avantage (Thermo VG Fisher), they are basically limited for the standalone XPS characterization utilized with the Mg and Al Ka anodes as a X-ray source, which generates a fixed photon energy of 1253.6 and 1486.6 eV, respectively. However, the most of traditional software are not proper to analyze the SR-based XPS data, because synchrotron radiation (SR) produces a wide spectrum and enables us to tune the photon energy. The monochromator and mirror optics in the beamline deliver SR to the spot on the sample surface in various energy resolutions and spot sizes. The electron energy and its intensity are measured by using the various electron energy analyzers and detectors. To visualize and compare a number of SR-based spectra, the scientific graphing and spread-sheet software such as Igor Pro, Origin, and MATLAB&trade; has been typically used to calibrate the photoelectron energy and intensity and normalize the photon flux and detection efficiency in each spectrum prior to the deepen chemical analysis. However, these software are relatively expensive for XPS beginners and further required for scripting in their own built-in macro languages to streamline spectral analyses from measurements to publications. The small-medium enterprises are also developing the sophisticated program such as AAnalyzer&reg; (www.rdataa.com), unifit (www.unifit-software.de), and KolXPD (www.kolibrik.net) for advanced ARPES, XPS, and NEXAFS analyses. However, the program cannot be shared with the collaborators without licenses. Common Data Processing System (www.sasj.jp/COMPRO/) works as a free program according to the ISO standard methods but its function is still limited to the standard XPS system.

Microsoft Excel is a de facto and standard spreadsheet to inspect and visualise the numerical data in various fields for daily jobs, because relational formulation among cells in spreadsheet and Solver non-linear optimisation function in Excel Add-In robustly handle scientific and engineering data as well as financial and accounting data. In addition, the VBA code streamlines the standard data analysis protocol without elaborated copy and paste actions on the spreadsheet, and instantly plots the charts in the worksheets. Even though the optimisation performance and numerical accuracy are quite limited in the big-data analysis, Excel VBA code makes your spectral data analysis simple and comprehensive on your own laptop. Python and R are robust and open-source scripting languages included with established analytical and statistical libraries, and their integrated development environment (IDE) is well supported and flexible for multiple purposes. However, the IDE setting is still a burden and complicated for beginners to inspect and evaluate the data quality and necessity for the further additional experiment during the limited beamtime. Excel worksheet is easily formulated and instantly distributed as a workbook for collaborators in academia and industries worldwide. Further analysis can be designed and performed in the Excel VBA environment under the terms of the GNU General Public License.

Ctrl+Q is a useful and flexible suite for SR-based XPS and XAS data analyses working on the Microsoft Excel VBA and its Solver Add-In, because Ctrl+Q streamlines the standard protocols of XPS and XAS data analyses in the four fundamental steps. First step is to calibrate the photon and electron energies, second to identify the peaks and its chemical shifts, third to subtract the spectral background and fit the peak, and fourth to evaluate the atomic concentration and ratio of chemical states. Ctrl+Q was initiated at the BL3.2Ua in the Siam Photon Laboratory operated by the Synchrotron Light Research Institute (Public Organization in Thailand), and has been optimized for user services since 2012. The beamline delivers the photo energy range from 40 to 1040 eV, so the most of function and database are also restricted in the range below 1,500 eV.

Any spectral data in the ASCII text included with energy and its corresponding intensity can be imported and formatted in the Excel spreadsheet. Supplementary information on the experimental conditions such as the beamline and detector settings are also described in the data file and extracted to calibrate and normalize spectra for quantitative analysis. The atomic elements in the sample can be identified from the peak energies based on the database of core level binding energy, which is typically referenced at the adventitious carbon 1s peak position of 284.6 eV in the binding energy scale. Auger electron peaks always appear in the XPS spectra at the constant kinetic energy, and their binding energies are varied with the photon energy. The database of the Auger electron enables us to distinguish between the Auger and photoelectron profiles. The amount of element can be evaluated in the ratio among the peak areas subtracted by the background and normalized with the atomic sensitivity factor, which is in fact varied with the photon energy to excite the photoelectron. The chemical shifts of the peaks, which correspond to the chemical states of the elements, can be disclosed in the peak fitting process optimized by the Excel Solver Add-In. The Shirley, Touggard, and polynomial background profiles and their combinations are available in either static and active optimization with peak fitting process. The Gaussian, Lorentzian, and pseudo-Voigt peak profiles are available with either asymmetric or tail parameters. The fitting parameters are easily constrained at the fixed value or limited range, and relative amplitude ratios and binding energy differences among the peaks are also controllable for doublet analysis.

XAS data analysis can also be performed in the same way as XPS with additional normalization and background subtraction processes. The pre-edge subtraction and post-edge normalization are available to subtract the absorption from the other elements and differentiate the atomic scattering factor, respectively. The arc-tangent and error functions are available to subtract the ionization background. The carbon K edge spectrum deformed by the carbon contamination in the beamline can be restored in the additional normalization with either the photon flux measured in the photo-diode or the carbon K edge spectrum of gold reference.

All the routine protocols implemented in the Ctrl+Q suite can be automated in the batching processes for multiple files to apply the same fitting parameters in the different spectra with the same series of samples. Each workbook contains the multiple worksheets for different purposes and functions to trace back the analysis sequences. For example, the original data is kept untouchable in the Data sheet. The Graph sheet produced from the Data sheet in the same workbook displays the plots of spectrum to inspect the peaks and chemical shifts. The Fit sheet generated from the data in the Graph sheet shows the fitting sequences to evaluate the peak shapes and areas. The Ana and Cmp sheets generated from the Fit sheets among the series of workbooks on a particular element disclose the ratios of chemical states in an element and background subtracted spectra. The Rto sheet generated from the Ana sheets among the series of workbooks among the series of workbooks includes the atomic concentration of elements. XAS spectral normalization, edge correction, and linear-combination fitting also add the Norm, Edge and Lcbn sheets, respectively.

The Ctrl+Q is simply triggered by the assigned short-cut key, and it functions different ways up to the active worksheet described above. For example, Ctrl+Q works on the Data sheet to generate the Graph sheet, on the Graph sheet to update the energy scale, and on the Fit sheet to optimize the fitting parameters. The specific syntax described in the cell on the worksheet with the short-cut key also triggers various functions.

Practical information and operation procedures are briefly described in the following sections.

## Installation
The VBA code is installed through VBE as the Personal Workbook Macro, and assigned with a short-cut key assignment in the macro called "PERSONAL.XLSB!CLAM2" listed at the top of sub procedures. Solver Add-in also needs to be deployed in Excel Add-ins and referenced in VBE to activate Solver Add-in for VBA code. Any data analysis sequence runs from the Option of Macro menu "PERSONAL.XLSB!CLAM2" or the short-cut key assigned on the worksheet in the workbook. The VBA code works differently on either what worksheet is active or what syntax is specified in active worksheet. See YouTube Video: https://youtu.be/tWpcnDjkHzo.

## Data loading
Any data placed in the Excel spreadsheet with specified formats can be analyzed in the VBA code. First of all, "KE/eV" is typed in the A1 cell for the spectrum with the kinetic energy scale. The kinetic energy data are in the first column below A2 and their corresponding intensities in the second column below B2. "BE/eV" at the A1 cell corresponds to the binding energy scale of spectrum, "PE/eV" the photon energy, and "ME/eV" any other scales. The name of workbook has to be careful because it is also used for the name of worksheet contained the spectral data called the Data sheet, and the other processing worksheets for graphing (Graph sheet), fitting (Fit sheet). The length of workbook name is also limited within 19 characters other than the extension characters after a dot, because the VBA code adds the initial identifiers (6 characters as maximum) of each worksheet name prior to the worksheet name of the Data sheet, even though the name of Excel worksheet is limited within 25 characters.

## Basic operation
The short-cut key is used to run the code on the Data sheet, and you can specify the photon energy and atomic elements in the dialogue boxes appeared during the processing. Eventually, the Graph and Fit sheets are generated from the Data sheet to tidy the data before the fitting process. Graph sheet is used to calibrate the energy and intensity, identify the peaks and chemical shifts, and compare the spectra reading from the other workbooks. The short-cut key is used to update the plots in the Graph and Fit sheets after the revisions of the parameters for calibrations and peak identification. Fit sheet is used to initiate the optimization of the background subtraction and peak fitting either in sequential or simultaneous. The parameters of fitting can be adjustable and constrained in the range after the first fitting sequence, and optimized repeatedly by using the short-cut key.

## Comparing spectra
You can compare the spectral data with one after another in the Graph and Fit sheets. In the Graph sheet, the short-cut key is used with syntax "comp" in the D1 cells to open the dialogue for selection of the workbooks to be compared. You can also add or replace the data one after another to use "comp" in every 3 columns after D1 cell like G1, J1, and so on. In the Fit sheet, the short-cut key is used with syntax "ana" in the D1 cell only to open the dialogue for selection of the workbooks to be compared, and produces the Cmp sheet. Compared spectra normalized and calibrated in Graph and Cmp sheets are easily exported into the Exp sheet by the short-cut key with syntax "exp" in A1 cell to be used for the other programs. 

## Energy calibration
Energy calibration is required to correct the energy scale for each spectrum, because the photon energy from the beamline is slightly deviated by the heat-load conditions and the electron energy from the sample is shifted by the charging or work function of instrument. It is difficult to evaluate the absolute photon and electron energies during the limited beamtime. However, the standard sample spectrum leads to the correct binding energy without evaluating the absolute energies, because the energies measured in the standard sample are well-known in the database or literature. It is also noted that the photon energy is used to evaluate the XPS sensitivity factors, so the roughly estimated photon energy is still required. In the Graph sheet, even though the photon energy, work function, and charging factor are adjustable parameters for binding energy correction, the calibration of binding energy is finalized in the charging factor. 

## Intensity scaling for inspection purpose
The offset and multiple factors are also available to subtract baseline and scale spectral intensity for data comparison. To compare the multiple spectra at a glance, both ends of spectral intensity are automatically subtracted to be zero and scaled to be unity with the short-cut ley with the syntax of "auto" at A1 cell in Graph sheet. To specify the energy ranges for spectral offset and multiple scaling, "auto[x0,x1:x2,x3]" can be used in a way that the spectral range between x0 and x1 is averaged to be 0 (offset) and the range between x2 and x3 is averaged to be 1 (multiple). If either x0/x2 or x1/x3 is null, nothing happens to be scaled in the corresponding range. The original data scales for offset and multiple factors are 0 and 1 respectively, and easily reset at these values by syntax "auto0". After the energy calibration, you can also repeat the previous scaling with syntax "autop". It should be noted that the fitting process is performed on the original data only.

## Spectral normalization
Spectral intensity is divided (normalized) by the other reference spectrum to compensate the noise or contamination happened during the measurement. Reference data can be added as the second data set with syntax "comp" in the Graph sheet prior to the normalization as mentioned above. The short-cut key with syntax "norm" at A1 cell in the Graph sheet proceeds the normalization of the first data set by the second data set resulting in the third data set in the Norm sheet. 

## Curve fitting
The peaks calibrated and identified by the database in the Graph sheet are analyzed in the Fit sheet based on the least-square regression method. Peak area is evaluated with analytical and numerical ways together with the choice of background subtraction processes. The number of peaks can be chosen with parameters such as curve shape, energy, FWHM width, amplitude etc. All the parameters can be constrained or limited in a specific range. Amplitude ratios and peak energy differences among the peaks are also adjustable in the cells with specified syntax described elsewhere. 

### Type of background subtraction and peak fitting function
- Gaussian, Lorentzian, and its blended function with tail parameters for asymmetric peak ("G" to "TSGL")
- Shirley and Tougaard backgrounds blended with and without polynomial function
- Constant, linear, quadratic, and cubic for polynomial background
- Polynomial and its blended backgrounds optimized with peak fittings (active mode: "BG" to "ABG")
- Arctangent and Victoreen backgrounds for XAS pre-edge subtraction
- Trapezoidal (numerical) integration for peak areas normalized by various sensitivity factors including photoionization cross section, source angle correction, MFP, analyzer transmission function etc.
- User-defined function can be easily implemented in the Visual Basic programming.

## Multiple data file analysis
Batching process can be customized from the macro called "PERSONAL.XLSB!debugAll".

## Notes
Ctrl+Q has been used for many users during the experiment and post-data processing to publish the data in the manuscript in peer-reviewed journals in the following.

- http://dx.doi.org/10.1016/j.snb.2013.12.017
- http://dx.doi.org/10.1016/j.apsusc.2016.01.180
- http://dx.doi.org/10.1016/j.carbon.2015.01.018
- http://dx.doi.org/10.1016/j.jenvman.2015.09.036
- http://dx.doi.org/10.1039/C6RA09972F

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

#### Supplementary information
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


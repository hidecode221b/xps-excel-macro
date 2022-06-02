
# Workflow
- Calc (XPS/XAS, PE, Elem)
    - Sim
    - Data (phi, CLAM2, KE/eV, BE/eV, PE/eV)
        - Check (CLAM2, exp)
            - Eck
        - Photo (TEY, TFY)
        - Graph (chem, elem, ana, exp, comp, auto, cali, noise, debug, norm, diff, cked, edge, lcmb, vms)
            - Exp
            - Norm
                - Graph_Norm
                    - Fit_Norm (exp)
                        - Exp_Fit_Norm
            - Diff
            - Edge
                - Graph_Edge
                    - Fit_Edge (exp)
                        - Exp_Fit_Edge
            - Lcmb
                - Graph_Lmcb
                    - Fit_Lmcb (exp)
                        - Exp_Fit_Lmcb
            - Fit (lmfit, ana, exp)
                - Pyt
                - Exp_Fit
                - Cmp
                - Ana
                    - Rto

# Data template
|Technique | Trigger in A1 | Queries | Graph x-axis |Fitting|
|:-----------|:------|:-------|:-------|:-------|
|PES|KE/eV|PE & elements|BE & KE|Yes in BE scale|
|XPS|BE/eV|PE & elements|BE & KE|Yes in BE scale|
|XAS|PE/eV|Elements|PE|Yes in PE scale|
|Grating scan|GE/eV|Gap/1st har.& e|PE|No|
|AES|AE/eV|Elements|EE & dN/dE|No|
|RGA|QE/eV|NA|Mas|Yes in mass|
|Manual scan|ME/eV|NA|Position|Yes in x|
|Histogram|HE/eV|NA|Position|Yes in x|
|Photodiode|FE/eV|Gap/1st har.|PE|No|


# Command
| Command | Cell | Sheet | Outcome |
|:-----------|:------|:-------|:-------|
|chem|C10|Graph, Cmp|display chemical shifts|
|elem|C10|Graph|input elements|
|intp|A1|Data|interpolate data by B1|
|ana|C10|Graph|update fit sheet|
|exp|A1|Graph, Check, Cmp|export data with E/eV name|
|expo|A1|Graph, Check, Cmp|export data for origin|
|expk|A1|Graph|export data with KE/eV in kinetic energy scale|
|comp|D1|Graph|compare data|
|auto|A1|Graph, Cmp|calibrate offset and multiple factors|
|cali|A1|Graph|calibrate C1s and Au4f|
|noise(n)|A1|Graph|denoise|
|ana|D4|Fit|summarize fit sheets|
|ana|A1|ana|summarize ana sheets to rto sheet|
|norm, diff|A1|Graph|normalize data|
|cked|A1|Graph|normalize data by gold C K data|
|edge|A1|Graph|edge correction|
|lmcb|A1|Graph|linear combination|
|vms|A1|Graph|export vamas (iso) format|
|phi|A2|Data|convert phi csv format to Excel|
|simulation|A1|Data|simulate spectrum|
|lmfit|D1|Fit|export python script for [lmfit](https://lmfit.github.io/lmfit-py/)|

# Background
| BG | A1 | B1 | C1 |
|:-----------|:------|:-------|:-------|
|[Shirley](https://doi.org/10.1103/PhysRevB.5.4709) BG|sh|ab/bg| |
|[Tougaard](https://doi.org/10.1002/sia.740110902) BG|to|bg| |
|Polynomial BG|po|ab/bg| |
|Polynomial Normal BG|po|no|ab/bg|
|Polynomial Shirley BG|po|sh|ab/bg|
|Polynomial Tougaard BG|po|to|bg|
|Polynomial Edge BG|po|ed|ab/bg|
|Polynomial AsLS BG|po|as|ab/bg|
|[Slope](https://doi.org/10.1016/j.elspec.2013.07.006) Shirley BG|sl|sh|ab/bg|
|Shirley Iterated BG|sh|it|bg|
|Shirley Peak BG|sh|pe|abg|
|[Arctan](https://doi.org/10.1103/RevModPhys.31.616) BG|ar|ab/bg||
|[Erf](https://doi.org/10.1063/1.453902) BG|er|ab/bg||
|[Victoreen](https://doi.org/10.1103/PhysRevB.11.4825) BG|vi|ab/bg||
|Double Exponential BG|do|ab/bg||
|Lognormal|lo|ab/bg||
|Sigmoid fit + spline BG|si|fi||
|Sigmoid convoluted fit|si|co|fi|
|Double Sigmoid fit|do|si|fit|
|User-defined function|ud|fit||
|SAXS|sa|fit||
|CK (C K edge on Arctan BG)|ck|||

# Peak shape (new version after 8.47)
| Syntax | Shape | G/L | Option a | Option b | #par|	Ref.|
|:-----------|:------|:------|:-------|:-------|:-------|:-------|
|G|Gaussian| 0 |||3||
|GL|Pseudo Voigt | 0-1 |||4|G + L with the same FWHM|
|TGL|Skewed Pseudo Voigt | 0-1 | skew ||5|G + L with the same FWHM|
|TG|Skewed Gaussian | 0 | skew ||4||
|EG|Exponential Gaussian | 0 | skew ||4||
|LN|Lognormal | 0 | skew ||4||
|L|Lorentzian | 1 |||3||
|SL|Split Lorentzian | 1 |||4| two FWHM|
|F|Breit Wigner Fano | 1 | skew ||4||
|DS|Doniac Sunjic| 1 | skew ||4||
|DSL|Doniac Sunjic Lorentzian | 1 | skew ||4||

# Peak shape (old version before 8.46)
| Syntax | Shape | Option a | Option b | #par|	Ref.|
|:-----------|:------|:-------|:-------|:-------|:-------|
|G (0)|Gaussian|||3||
|*DB G (0)*|Double Gaussian|||4|[Fityk](https://fityk.nieto.pl/)|
|_*EMG*_|Exponentially Modified Gaussian|Distortion para.||4|[Fityk](https://fityk.nieto.pl/)|
|L (1)|Lorentzian|||3||
|*DS L (1)*|Doniac-Sunjic x L|Asymmetric para.||5|[CasaXPS](http://www.casaxps.com/)|
|_*DB L (1)*_|Double Lorentzian|||4|[AAnalyzer](http://rdataa.com/aanalyzer/aanaHome.htm)|
|_PEA_|Pearson VII|Skewness||4|[Fityk](https://fityk.nieto.pl/)|
|SGL, PGL (0-1)|G + L, G x L (pseudo-Voigt)|||5|[Unifit](https://www.unifit-software.de/),[CasaXPS](http://www.casaxps.com/)|
|*ASGL, APGL*|Asymmetric V, Double Voigt|||5|[doi](https://doi.org/10.1107/S0021889884011043)|
|ESGL, EPGL|Exponential blended Voigt|Exponential decay parameters||5|[CasaXPS](http://www.casaxps.com/)|
|_*DS SGL, DS PGL*_|DS x L blended V|Asymmetric parameter|Ratio DSL:V|6|[CasaXPS](http://www.casaxps.com/)|
|\_UG SGL, UG PGL_|Ulrik Gelius blended Voigt|Asymmetric parameter a|Asymmetric parameter b|6|[CasaXPS](http://www.casaxps.com/)|
|\_*DSV SGL, DSV PGL*_|	DS x Voigt blended Voigt|Asymmetric parameter|Ratio DSV:V|6|[CasaXPS](http://www.casaxps.com/)|
|\_TSGL_| 	Exponential blend SGL (MultiPak) |Tail scale| Tail length at half max| 6|[MultiPak](https://www.ulvac-phi.com/)|
|GL (0 < shape < 1) |G + L with the same FWHM (MultiPak) |||4|[MultiPak](https://www.ulvac-phi.com/), Eq. to SGL|
|MSGL|Asymmetric Voigt|Asymmetric parameter|Sigmoid-center translation|6|[doi](https://doi.org/10.1002/sia.5521)|
|CGL|Numerical convolution G x L|||4|[doi](https://doi.org/10.1002/sia.2527)|
|F|	Fano profile|||4|[doi](https://doi.org/10.1103/PhysRev.124.1866)|
|FG|F x G|||5||
|LOGN|Log normal|Mean (μ)||4||

# Optimization mode of fittings
| Cell in Fit sheet | Syntax or Font style | Optimization|
|:-----------|:------|:-------|
|BE, FWHM, Ampl, Shape, Options|Figures with Bold|Constraints|
|A14|Solve chi^2*|Least chi square|
|A14|Solve Abbe|Abbe criteria |
|A10 (EF fit)|Solve FD without Italic|Least chi square|
|A10 (EF fit)|Solve FD with Italic|Abbe criteria |
|A11 (EF fit) |Solve GC without bold|Gaussian convolution after FD + polynomial BG|
|A11 (EF fit)|Solve GC with bold|FD + Polynomial BG first, Gaussian convolution together with FD + poly BG|


# Calibrations in offset/multiple factors
|A1 cell syntax in Graph sheet| Offset factor | Multiple factor|
|:-----------|:------|:-------|
|auto0	|Set to 0	|Set to 1|
|auto or auto1	|First point to be zero 	|End point to be unity |
|auto10 	|Zero at point 10 from start point 	|Unity at point 10 from end point |
|auto(1,10) 	|Zero from point 1 to 10 from start point 	|Unity from point 1 to 10 from end point |
|auto[100:101,200:201] 	|Zero in BE range between 100 and 101 eV 	|Unity in BE range between 200 and 201 eV |
|automax / autowf	|Zero at the lower side of a point of data	|Unity at max intensity point of data|
|autop	|Syntax previously done	|Syntax previously done|
|auto{284.6}	|BE at max. intensity to be calibrated in 284.6 eV	|NA (BE calibration by Charging factor)|
|auto'-7.8'	|Charging correction at -7.8 eV for all spectra	|NA (this is based on C1s BE calibration)|
|offset10	|Offset spectra for water fall plot	|NA|

# List of element groups to be identified
|Syntax|Group|Elements to be analyzed|
|:-----------|:------|:-------|
|AL|Alkali metals|Na,K,Rb,Cs|
|EA|Alkaline Earth metals|Be,Mg,Ca,Sr,Ba,Ra|
|TM|Transition metals|3d + 4d + 5d transition metals|
|3d|3d transition metals|Sc,Ti,V,Cr,Mn,Fe,Co,Ni,Cu,Zn|
|4d|4d transition metals|Y,Zr,Nb,Mo,Tc,Ru,Rh,Pd,Ag,Cd|
|5f|5d transition metals|Lu,Hf,Ta,W,Re,Os,Ir,Pt,Au,Hg|
|SM|Semi-metals|B,Si,Ge,As,Sb,Te|
|NM|Non-metals|C,N,O,P,S,Se|
|BM|Basic metals|Al,Ga,In,Sn,Tl,Pb,Bi|
|HA|Halogens|F,Cl,Br,I,At|
|NG|Noble gases|Ne,Ar,Kr,Xe,Rn|
|RM|Rare metals|La,Ce,Nd,Sm,Eu,Gd,Tb,Er,Tm,Yb,Th,U|
|LA|Lanthanide|La,Ce,Nd,Sm,Eu,Gd,Tb,Er,Tm,Yb|
|AC|Actinides|Th,U|


# Advanced syntax templates in the sheets 
|Purpose|Sheet|Cells|Formula| Reference cell|Calibrated #1|Calibrated #2|
|:-----------|:------|:-------|:-----------|:------|:-------|:-------|
|Extra photons|Graph|C2|;100;200;333 eV||||		
|Specific scans|Graph|B8|[1,2-4]||||	
|Amp ratio|Fit|D14-|(4;1;3)|(4;|1;|3)|
|BE diff|Fit|D15-|[3.5;n3.5]|[|3.5;|n3.5]|

Note1: “n” represents negative shift from reference.

Note2: Empty cells between brackets does not effect to the constraints.

# List of Peak area 
|Name|Usage|Description|Factors to be effective|
|:-----------|:------|:-------|:-------|
|P. Area|Chemical state analysis|Peak area calculated with analytical formula and without any factors|Amplitude, FWHM|
|S. Area|Quantification of elements under the same condition|Peak area normalized with atomic sensitivity factor based on photo-ionization cross-section|Amplitude, FWHM, PE, Sensitivity based on element specified in the Graph sheet|
|N. Area|Quantification of elements under the various measurement conditions|Peak area calculated in "S. Area" plus normalized with empirically calculated factors at BL CLAM2 including XPS mean-free path of photoelectrons, transmission function of electron energy analyzer based on pass energy, grating efficiency|Amplitude, FWHM, PE, KE, Sensitivity, CAE, Grating, MFP factor, a & b specified in the Fit sheet based on formalism from CasaXPS|

T.I./S.I./N.I. are numerically integrated areas with Trapezoidal rule applied to each corresponding area shown above.

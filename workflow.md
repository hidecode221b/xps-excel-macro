
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
                    
# Syntax
| Command | Cell | Sheet | Outcome |
|:-----------|:------|:-------|:-------|
|chem|C10|Graph, Cmp|display chemical shifts|
|elem|C10|Graph|input elements|
|intp|A1|Data|interpolate data by B1|
|ana|C10|Graph|update fit sheet|
|exp|A1|Graph, Check, Cmp|export data with unique name|
|exp2|A1|Graph, Check, Cmp|export data with E/eV name|
|exp3|A1|Graph|export data with AE/eV for Auger|
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
|vms|A1|Graph|export vamas format|
|phi|A2|Data|convert phi csv format to Excel|
|simulation|A1|Data|simulate spectrum|
|lmfit|D1|Fit|export python script for lmfit|

# Background
| BG | A1 | B1 | C1 |
|:-----------|:------|:-------|:-------|
|Shirley BG|sh|ab/bg| |
|Tougaard BG|to|ab/bg| |
|Polynomial BG|po|ab/bg| |
|Polynomial Normal BG|po|no|ab/bg|
|Polynomial Shirley BG|po|sh|ab/bg|
|Polynomial Tougaard BG|po|to|ab/bg|
|Polynomial Conv-Tougaard|po|co|ab/bg|
|Polynomial Virtual Shirley BG|po|vi|ab/bg|
|Polynomial Edge BG|po|ed|ab/bg|
|Polynomial AsLS BG|po|as|ab/bg|
|Slope Shirley BG|sl|sh|ab/bg|
|Slope Tougaard BG|sl|to|ab/bg|
|Slope Virtual Shirley BG|sl|vi|ab/bg|
|Shirley Iterated BG|sh|it|bg|
|Shirley Peak BG|sh|pe|abg|
|Virtual Shirley BG|vi|sh|ab/bg|
|Tougaard Convoluted|to|co|ab/bg|
|Arctan BG|ar|ab/bg||
|Erf BG|er|ab/bg||
|Victoreen BG|vi|ab/bg||
|Double Exponential BG|do|ab/bg||
|Lognormal|lo|ab/bg||
|Sigmoid fit + spline BG|si|fi||
|Sigmoid convoluted fit|si|co|fi|
|Double Sigmoid fit|do|si|fit|
|User-defined function|ud|fit||
|SAXS|sa|fit||
|CK (C K edge on Arctan BG)|ck|||

# Peak shape
| Syntax | Shape | Option a | Option b | #par|	Ref.|
|:-----------|:------|:-------|:-------|:-------|:-------|
|G (0)|Gaussian|||3||
|DB G (0)|Double Gaussian|||4|Fityk|
|EMG|Exponentially Modified Gaussian|Distortion para.||4|Fityk|
|L (1)|Lorentzian|||3||
|DS L (1)|Doniac-Sunjic x L|Asymmetric para.||5|CasaXPS|
|DB L (1)|Double Lorentzian|||4|AAnalyzer|
|PEA|Pearson VII|Skewness||4|Fityk|
|SGL, PGL (0-1)|G + L, G x L (pseudo-Voigt)|||5|Unifit,CasaXPS|
|ASGL, APGL|Asymmetric V, Double Voigt|||5|10.1107/S0021889884011043|
|ESGL, EPGL|Exponential blended Voigt|Exponential decay parameters|||5|CasaXPS|
|DS SGL, DS PGL|DS x L blended V|Asymmetric parameter|Ratio DSL:V|6|CasaXPS|
|UG SGL, UG PGL|Ulrik Gelius blended Voigt|Asymmetric parameter a|Asymmetric parameter b|6|CasaXPS|
|DSV SGL, DSV PGL|	DS x Voigt blended Voigt|Asymmetric parameter|Ratio DSV:V|6|CasaXPS|
|TSGL| 	Exponential blend SGL (MultiPak) |Tail scale| Tail length at half max| 6|MultiPak|
|GL (0 < shape < 1) |G + L with the same FWHM (MultiPak) |||4|MultiPak, Eq. to SGL|
|MSGL|Asymmetric Voigt|Asymmetric parameter|Sigmoid-center translation|6|10.1002/sia.5521|
|CGL|Numerical convolution G x L|||4|10.1002/sia.2527|
|F|	Fano profile|||4|10.1103/PhysRev.124.1866|
|FG|F x G|||5||
|LOGN|Log normal|Mean (μ)||4||


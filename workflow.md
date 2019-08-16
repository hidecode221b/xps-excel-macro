
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

# Fitting
| BG | A1 | B1 | C1 |
|:-----------|:------|:-------|:-------|
|Shirley BG|sh|ab/bg| |
|Tougaard BG|to|ab/bg| |
|Polynomial BG|po|ab/bg| |

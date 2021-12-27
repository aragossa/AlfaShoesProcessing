def get_excel_formula(col_num, row_num):
    excel_formulas = {
        15: f'=IF(VLOOKUP($M{row_num},stw_modules_models_local,1,1)<>$M{row_num},"В системе нет такой модели!",IF(VLOOKUP($M{row_num},stw_modules_models_local,6,1)<>0,"Указать колодку",""))',
        17: f'=IF(VLOOKUP($M{row_num},stw_modules_models_local,1,1)<>$M{row_num},"В системе нет такой модели!","A"&VLOOKUP($M{row_num},stw_modules_models_local,4,1)&IF(VLOOKUP($K{row_num},stw_modules_podklads_local,5,1)="нет","",VLOOKUP($K{row_num},stw_modules_podklads_local,5,1))&"-"&IF($O{row_num}="",VLOOKUP($M{row_num},stw_modules_models_local,5,1),$P{row_num}))',
        20: f'=IF(VLOOKUP($J{row_num},stw_modules_colors_local,1,1)<>$J{row_num},"В системе нет такой кожи!",VLOOKUP($J{row_num},stw_modules_colors_local,5,1))',
        24: f'=IF(VLOOKUP($I{row_num},stw_modules_sizes_local,1,1)<>$I{row_num},"В системе нет такого размера!",VLOOKUP($I{row_num},stw_modules_sizes_local,3,1))',
        26: f'=IF($AB{row_num}="",(IF($R{row_num}="",$Q{row_num},$R{row_num})&"-"&IF($U{row_num}="",$T{row_num},$U{row_num})),"")',
        27: f'=IF($AB{row_num}="",($Z{row_num}&"|"&IF($Y{row_num}="",$X{row_num},$Y{row_num})),($AB{row_num}&"|"&IF($R{row_num}="",$Q{row_num},$R{row_num})&"-"&IF($U{row_num}="",$T{row_num},$U{row_num})&"|"&IF($Y{row_num}="",$X{row_num},$Y{row_num})))',
        28: f'=IF($N{row_num}="не серия",$B{row_num},"")',
        29: f'=IF($R{row_num}="",$Q{row_num},$R{row_num})',
        30: f'=IF($U{row_num}="",$T{row_num},$U{row_num})',
        31: f'=IF($Y{row_num}="",$X{row_num},$Y{row_num})',
        32: f'=VLOOKUP($K{row_num},stw_modules_podklads_local,4,1)',
        33: f'=IF($W{row_num}="",VLOOKUP($AD{row_num},кожа,4,FALSE),$W2)',
        35: f'=IF($F{row_num}="удален",1,IF($L{row_num}=1,1,0))',
        36: f'=IF(OR($F{row_num}=1,$F{row_num}=2,$F{row_num}=3,$F{row_num}=4),1,0)',
        37: f'=IF(OR($F{row_num}=5,$F{row_num}=6),1,0)',
        38: f'=IF(AND($F{row_num}=7,$L{row_num}=0),1,0)',

        39: f"=IFERROR(VLOOKUP($A{row_num},сдано_на_склад,'Сдано на склад'!$D$1,FALSE),0)",
        40: f"=IFERROR(VLOOKUP($A{row_num},сдано_на_склад,'Сдано на склад'!$E$1,FALSE),"")",
        41: f"=IFERROR(VLOOKUP($A{row_num},сдано_на_склад,'Сдано на склад'!$I$1,FALSE),"")",
        42: f"=IFERROR(VLOOKUP($A{row_num},сдано_на_склад,'Сдано на склад'!$S$1,FALSE),"")",
        43: f"=IF(OR($AJ{row_num}=1,$AK{row_num}=1,$AL{row_num}=1),IF($AM{row_num}=1,0,1),0)",

        46: f'=IF($F{row_num}="удален","удален",IF(AND($AR{row_num}="",$AS{row_num}="",$E{row_num}=0,$H{row_num}=0),"",($AR{row_num}&" "&$AS{row_num}&" "&IF($E{row_num}=0,"",$E{row_num})&" "&IF($H{row_num}=0,"",$H{row_num}))))',
        47: f'=IF($F{row_num}="удален","удален",IF(LEN($A{row_num})=3,"200000000"&$A{row_num},IF(LEN($A{row_num})=4,"20000000"&$A{row_num},IF(LEN($A{row_num})=5,"2000000"&$A{row_num},IF(LEN($A{row_num})=6,"200000"&$A{row_num},0)))))',
        48: '=IF($F%s="удален","удален",MOD(10-MOD(SUM(MID($AU%s,{1,2,3,4,5,6,7,8,9,10,11,12},1)*{1,3,1,3,1,3,1,3,1,3,1,3}),10),10))' % (
        row_num, row_num),
        49: f'=IF($F{row_num}="удален","удален",$AU{row_num}&$AV{row_num})',
        51: f'=IF($F{row_num}="удален","Изделие удалено",IF(VLOOKUP($A{row_num},актуальная_выгрузка,1,1)<>$A{row_num},"",IF(OR(VLOOKUP($A{row_num},актуальная_выгрузка,20,1)<>"",VLOOKUP($A{row_num},актуальная_выгрузка,21,1)<>"",VLOOKUP($A{row_num},актуальная_выгрузка,22,1)<>"",VLOOKUP($A{row_num},актуальная_выгрузка,23,1)<>"",VLOOKUP($A{row_num},актуальная_выгрузка,24,1)<>"",VLOOKUP($A{row_num},актуальная_выгрузка,25,1)<>""),VLOOKUP($A{row_num},актуальная_выгрузка,20,1)&" "&VLOOKUP($A{row_num},актуальная_выгрузка,21,1)&" "&VLOOKUP($A{row_num},актуальная_выгрузка,22,1)&" "&VLOOKUP($A{row_num},актуальная_выгрузка,23,1)&" "&VLOOKUP($A{row_num},актуальная_выгрузка,24,1)&" "&VLOOKUP($A{row_num},актуальная_выгрузка,25,1),"")))'
    }
    return excel_formulas.get(col_num)

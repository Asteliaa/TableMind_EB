"""Сохранение единичных рядов GT (по одному запросу) и 12-месячных пакетов."""
from pathlib import Path

OUT = Path(__file__).resolve().parent.parent / "Материалы_собранные" / "GT"
P1 = "excel formula checker ; spreadsheet audit ; excel error checker ; excel model audit ; check excel formulas"
P3 = "excel copilot ; chatgpt excel ; excel ai ; perfectxl ; claude for excel"
RU1 = "проверка формул excel ; ошибки в формулах excel ; аудит excel ; нейросеть для excel ; excel искусственный интеллект"
# id, geo, time, start_epoch, n, keywords, checksum, data
ITEMS = [
    ("S_excel_formula_checker_5y", "WORLD", "today 5-y", 1632614400, 262, "excel formula checker", "3012",
     "~8_c~7_W00_300r~4n00_3~7_B~G_4x~Av00_50v_3_M~6r_4~3w~Fx00_G0s0_M0_H_V_K_I~6_4~5_A00_9~4_7~C_9~8z_800_6~p_0_50_N~F_00_2~3_100w00_Ey0_M~B6JI"),
    ("S_spreadsheet_audit_5y", "WORLD", "today 5-y", 1632614400, 262, "spreadsheet audit", "2319",
     "~A6~6607005C709~3956~755~56650A~39050500605074005~5500555060600666~A75500905~460060050606~37069680597006~36~3605060A6967700700909007006~47~3666~48~3867790998EKKHGBDGBAAFAGGPMMFLA9GIJPOLbZfgVZkgVR_P_c_RRSbbrxUcfSdHWOKaepWF"),
    ("S_excel_audit_5y", "WORLD", "today 5-y", 1632614400, 262, "excel audit", "4972",
     "88698799789987CDAABFCCCCAABC9ABAAA98ABAAAA9DAABCABDEFBBCABBB9BBBA8BCDBC9ABBBDCDDBBAABBBBB9AACCACCBBDCCCCEBBCAB999B9AA7ABCBDDEBBDDEECCBABACDBBCAECDEECBBDADBFGDCCEAAAB9CCB78BBDBCCDDBBDBDBACCCCCCFOIIFFFHINQQPQPUXROOSTUSWWYSSHGNQRTUVZeocbiz_3to_D_Z_cy_2_J_F_M_Fw_2lj_2ekk_9_5_9_7eK"),
    ("S_financial_model_audit_5y", "WORLD", "today 5-y", 1632614400, 262, "financial model audit", "1804",
     "~53~D3302223200220203~43~A3022~322~B2002003002002~42~J2~E304~4202~63~623~62033330232~34~7330320433~333~4476545465DGDBCCAGDHCCBFEHKJELCBDFKIJKPhfQVSguXS_8_c_XYZjaqzUWIMNGGENUWYPH"),
    ("S_excel_copilot_5y", "WORLD", "today 5-y", 1632614400, 262, "excel copilot", "5560",
     "~z~61~6107C765534333333333322222222325434354566664357CEDDFGFFEGEEEEEDCDFEHGDDBCBBBBBDCCFHIGHGIGIIKHHLHCDHMSTRPTRPPSRPQNSOPSYWccbWXXYVWht_c_6uuw_5_X_B_1t_1st_4rx_0uVUnoosuxu_2_5_K_H_8_3x_A_H_C_H_X_U_F_I_F_8_8_2_3pwnlirkpqsro"),
    ("S_perfectxl_5y", "WORLD", "today 5-y", 1632614400, 262, "perfectxl", "121", "~z~S_c~z~z~mL0"),
    ("W_P1_12m", "WORLD", "today 12-m", 1759017600, 53, P1, "60 1662 5 495 608",
     "0B0790A07D0A00A0F0995A0886G08G0G0AC7P07B0M0BA0M08P0F06B0L0980A~490070G00C0I0790J00D0P0660O06E0L0GD0b08E0Z07C0f0FD5g0FH0V0EC5Z0FB0k0CD0g0DA0V59E5R0CD0_P0HH0_c0MK5_R0SH0R0KF0S0HE6b0DD5b0HE0r0HF7x0GE0U0CA0c0HC0f0680S00D0d08C0H0590W0060O0080K06C0a0880e06A0p0922W0482F026"),
    ("W_P3_12m", "WORLD", "today 12-m", 1759017600, 53, P3, "855 1416 3357 0 255",
     "PX_802JY_001GVy01EWz01GVv04EX_604EW_102HV_201EU_203FW_E03GW_402FQv028Ia018IW02DPs02DRs03DQv03ERp05FRt06FQ_107FO_108HQ_A08IU_I09LV_D09LQ_E09IS_H09HP_G08FR_F08JP_L09LP_D08JP_G07LR_V09PV_a09OV_K06KU_M06LZ_c07KU_S07IT_P06IR_L06HQ_K06HQ_A06DNn04FQn04DRr04CNc03CMX03ELb03CJY03DMf03ENh03ELi04EGW03DIR03"),
    ("RU_R1_12m", "RU", "today 12-m", 1759017600, 53, RU1, "0 0 0 100 0", "~z~z~z~z~E_c~6"),
]
for id_, geo, time, start, n, kws, cs, data in ITEMS:
    txt = (
        "source Google Trends (explore API, web search, category 0)\nfetched 2026-10-02\n"
        f"id {id_}\ngeo {geo}\ntime {time}\nstart_epoch {start}\nn {n}\nstep_sec 604800\n"
        f"keywords {kws}\nchecksum {cs}\ndata {data}\n"
    )
    (OUT / f"{id_}.txt").write_text(txt, encoding="utf-8")
print("saved", len(ITEMS))

"""Сохранение небольших выгрузок GT (сжатый формат) в Материалы_собранные/GT."""
from pathlib import Path

OUT = Path(__file__).resolve().parent.parent / "Материалы_собранные" / "GT"
EN1 = "excel formula checker ; spreadsheet audit ; excel error checker ; excel model audit ; check excel formulas"
RU1 = "проверка формул excel ; ошибки в формулах excel ; аудит excel ; нейросеть для excel ; excel искусственный интеллект"
ITEMS = [
    ("US_P1_5y", "US", EN1, "186 1795 64 57 1472",
     "~9V~z~NX~ON~9O0L~MK~dR~z~DP~9N~dT~OR~OQ~FN~8N~9S~dS~z~IW~EY~YO~YS~4b~4a~9W~Ee~5TU~3d~aV~aX~kg~CY~9e~9Y~Eb~iQ~4V~4T~QS~z~QS~Be00c~BV~4b~9U00e0V00Y0R~EU00R00X0U~4T~5T~AV0RU~6X~3X~FQ00Y~6f~4f00a0_700c0o~4k00X0Z~4Y~4d~4k~4g00R0_6~4n~4b~4_E~4_c00f0_M~9V~4s~4_6~4_400S0_G~4_800S0v0RV0b~hX40020K00B0G03D"),
    ("GB_P1_5y", "GB", EN1, "224 34 0 6 76", "~Y_5~5_6~z~3u~z~z~z~z~z~z~z~z~z~z~z~z~z~V_c~z~z~z~z~z~z~4K0090E060"),
    ("DE_P1_5y", "DE", EN1, "0 141 160 63 77", "~z~z~z~a_F~z~z~z~z~z~Wy~K_1~z~z~z~z~z~r_c~z~z~B_4~s_D~z~z~z"),
    ("FR_P1_5y", "FR", EN1, "0 68 100 0 71", "~z~h_9~z~z~z~z~z~z~z~J_6~z~z~z~z~z~z~z~z~z~z~z~i_c~g"),
    ("NL_P1_5y", "NL", EN1, "0 168 0 198 159", "~z~z~z~z~z~Q_H~z~G_W~z~z~H_3~z~z~z~z~z~z~C_c~x_a~z~z~z~z~c_R~c"),
    ("BY_R1_5y", "BY", RU1, "94 241 0 165 0", "~z~z~x_H~z~z~z~D_3~z~z~Y_W~z~s_H~z~z~T_c~H_L~z~z~z~z~z~z~z~z"),
    ("RU_R1_5y", "RU", RU1, "0 100 87 102 65", "~w_V~z~z~z~gx~V_c~z~z~z~2_P~z~z~z~z~z~z~z~z~z~z~z~z~z~79~56"),
    ("KZ_R1_5y", "KZ", RU1, "0 100 86 0 117", "~z~z~z~z~z~z~Xk~z~z~v_9~z~z~z~z~z~z~z~9_c~z~z~z~b_O~z~B"),
]
for id_, geo, kws, cs, data in ITEMS:
    txt = (
        "source Google Trends (explore API, web search, category 0)\nfetched 2026-10-02\n"
        f"id {id_}\ngeo {geo}\ntime today 5-y\nstart_epoch 1632614400\nn 262\nstep_sec 604800\n"
        f"keywords {kws}\nchecksum {cs}\ndata {data}\n"
    )
    (OUT / f"{id_}.txt").write_text(txt, encoding="utf-8")
print("saved", len(ITEMS))

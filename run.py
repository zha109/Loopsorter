from dlssp_pipeline import run_dlssp_example, run_pipeline_from_excel
import pandas as pd

USE_EXAMPLE = False
EXCEL_PATH = r"C:\Users\Admin\Downloads\lssp_louvain_sample.xlsx"

if USE_EXAMPLE:
    print("=== RUNNING DLSSP EXAMPLE ===")
    result = run_dlssp_example()
else:
    print(f"=== RUNNING DLSSP PIPELINE FROM EXCEL: {EXCEL_PATH} ===")
    result = run_pipeline_from_excel(EXCEL_PATH)

out_path = EXCEL_PATH.replace(".xlsx", "_result.xlsx") if not USE_EXAMPLE else "dlssp_example_result.xlsx"

def _part_to_df(part: dict, label="cid"):
    return pd.DataFrame({"order": list(part.keys()), label: list(part.values())}).sort_values(["order"])

def _groups_to_df(groups: dict):
    rows = []
    for cid, members in groups.items():
        for o in members:
            rows.append({"cid": cid, "order": o})
    return pd.DataFrame(rows).sort_values(["cid", "order"])

with pd.ExcelWriter(out_path, engine="xlsxwriter") as w:
    M = result["G"].nodes()
    _part_to_df(result["p1_part"], "cid_p1").to_excel(w, sheet_name="Partition_P1", index=False)
    _part_to_df(result["p2_part"], "cid_p2").to_excel(w, sheet_name="Partition_P2", index=False)
    _part_to_df(result["part_final"], "cid_final").to_excel(w, sheet_name="Partition_Final", index=False)
    _groups_to_df(result["groups_final"]).to_excel(w, sheet_name="Groups_Final", index=False)
    df_cinfo = pd.DataFrame([
        {
            "cluster_id": cid,
            "size": len(ci["members"]),
            "load": ci["load"],
            "center": ci["center"],
            "express_ratio": ci["express_ratio"],
            "members": ", ".join(ci["members"])
        } for cid, ci in result["cluster_info"].items()
    ]).sort_values("cluster_id")
    df_cinfo.to_excel(w, sheet_name="Cluster_Info", index=False)

print(f"\n✓ DONE. Results saved to: {out_path}")

import pandas as pd

from excel_dashboard import prepare_dashboard_data
from storage.lineage import LINEAGE_COLUMNS


def build_slik_dwh_df(facilities_df, clean_df=None):
    if facilities_df is None:
        return pd.DataFrame()

    if clean_df is None:
        clean_df = prepare_dashboard_data(facilities_df)

    if clean_df is None or clean_df.empty:
        return pd.DataFrame(columns=list(getattr(clean_df, "columns", [])) + LINEAGE_COLUMNS)

    lineage_columns = [column for column in LINEAGE_COLUMNS if column in facilities_df.columns]
    if not lineage_columns:
        return clean_df.reset_index(drop=True)

    lineage_df = facilities_df.reindex(clean_df.index)[lineage_columns].reset_index(drop=True)
    return pd.concat([clean_df.reset_index(drop=True), lineage_df], axis=1)


def build_slik_marts(dwh_df):
    if dwh_df is None or dwh_df.empty:
        return {
            "mart_slik_summary": pd.DataFrame(),
            "mart_slik_risk": pd.DataFrame(),
            "mart_slik_maturity": pd.DataFrame(),
        }

    df = dwh_df.copy()
    for column in ["Plafond (Rp)", "Baki Debet (Rp)", "Rate"]:
        if column in df.columns:
            df[column] = pd.to_numeric(df[column], errors="coerce").fillna(0)

    summary = pd.DataFrame(
        [
            {
                "facility_count": int(len(df)),
                "debtor_count": int(df["Nama"].nunique()) if "Nama" in df.columns else 0,
                "bank_count": int(df["Bank"].nunique()) if "Bank" in df.columns else 0,
                "total_plafond": df["Plafond (Rp)"].sum() if "Plafond (Rp)" in df.columns else 0,
                "total_baki_debet": df["Baki Debet (Rp)"].sum() if "Baki Debet (Rp)" in df.columns else 0,
                "average_rate": df["Rate"].mean() if "Rate" in df.columns and len(df) else 0,
                "source_file_count": int(df["source_file"].nunique()) if "source_file" in df.columns else 0,
                "job_id": df["job_id"].dropna().iloc[0] if "job_id" in df.columns and not df["job_id"].dropna().empty else None,
            }
        ]
    )

    risk_group_columns = [column for column in ["Keterangan", "Kol."] if column in df.columns]
    if risk_group_columns:
        risk = (
            df.groupby(risk_group_columns, dropna=False)
            .agg(
                facility_count=("Nama", "size"),
                total_baki_debet=("Baki Debet (Rp)", "sum"),
                total_plafond=("Plafond (Rp)", "sum"),
            )
            .reset_index()
        )
    else:
        risk = pd.DataFrame()

    if "Maturity Bucket" in df.columns:
        maturity = (
            df.groupby("Maturity Bucket", dropna=False)
            .agg(
                facility_count=("Nama", "size"),
                total_baki_debet=("Baki Debet (Rp)", "sum"),
                total_plafond=("Plafond (Rp)", "sum"),
            )
            .reset_index()
        )
    else:
        maturity = pd.DataFrame()

    return {
        "mart_slik_summary": summary,
        "mart_slik_risk": risk,
        "mart_slik_maturity": maturity,
    }

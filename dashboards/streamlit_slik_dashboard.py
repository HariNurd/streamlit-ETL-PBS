import pandas as pd
import streamlit as st

try:
    import plotly.express as px
except ImportError:  # pragma: no cover - Streamlit fallback for lean installs.
    px = None


BUCKET_ORDER = ["Overdue", "0-3 Months", "3-6 Months", "6-12 Months", ">12 Months", "Unknown"]
BUCKET_LABELS = {
    "Overdue": "Lewat Jatuh Tempo",
    "0-3 Months": "0-3 Bulan",
    "3-6 Months": "3-6 Bulan",
    "6-12 Months": "6-12 Bulan",
    ">12 Months": ">12 Bulan",
    "Unknown": "Tidak Diketahui",
}


def _format_rupiah(value):
    value = float(value or 0)
    return "Rp " + f"{value:,.0f}".replace(",", ".")


def _format_percent(value):
    if pd.isna(value):
        return "-"
    return f"{float(value) * 100:.2f}%"


def _options(df, column):
    if column not in df.columns or df.empty:
        return []
    values = df[column].dropna().astype(str).str.strip()
    return sorted(value for value in values.unique() if value)


def _filter_by_values(df, column, selected):
    if not selected or column not in df.columns:
        return pd.Series(True, index=df.index)
    return df[column].fillna("").astype(str).isin(selected)


def _show_bar(data, x_col, y_col, title):
    if data.empty:
        st.caption("Tidak ada data untuk grafik ini.")
        return

    if px:
        fig = px.bar(data, x=x_col, y=y_col, title=title)
        fig.update_layout(margin=dict(l=10, r=10, t=48, b=10), xaxis_title=None, yaxis_title=None)
        st.plotly_chart(fig, use_container_width=True)
        return

    st.caption(title)
    st.bar_chart(data.set_index(x_col)[y_col])


def _show_histogram(data, value_col, title):
    if data.empty:
        st.caption("Tidak ada data untuk grafik ini.")
        return

    if px:
        fig = px.histogram(data, x=value_col, nbins=20, title=title)
        fig.update_layout(margin=dict(l=10, r=10, t=48, b=10), xaxis_title=None, yaxis_title=None)
        st.plotly_chart(fig, use_container_width=True)
        return

    st.caption(title)
    st.bar_chart(data[value_col])


def _metric_grid(metrics, columns=4):
    for start in range(0, len(metrics), columns):
        cols = st.columns(columns)
        for col, (label, value) in zip(cols, metrics[start:start + columns]):
            col.metric(label, value)


def render_slik_dashboard(clean_df, errors_df=None):
    st.subheader("Dashboard Preview SLIK")

    if clean_df is None or clean_df.empty:
        st.info("Tidak ada data fasilitas untuk ditampilkan di dashboard.")
        if errors_df is not None and not errors_df.empty:
            st.subheader("File Error")
            st.dataframe(errors_df, use_container_width=True)
        return

    with st.expander("Filter Dashboard", expanded=True):
        c1, c2, c3, c4 = st.columns(4)
        selected_nama = c1.multiselect("Nama", _options(clean_df, "Nama"), key="slik_filter_nama")
        selected_nik = c2.multiselect("NIK/NPWP", _options(clean_df, "NIK/NPWP"), key="slik_filter_nik")
        selected_bank = c3.multiselect("Bank", _options(clean_df, "Bank"), key="slik_filter_bank")
        selected_kol = c4.multiselect("Kol.", _options(clean_df, "Kol."), key="slik_filter_kol")

        c5, c6 = st.columns(2)
        selected_keterangan = c5.multiselect(
            "Keterangan",
            _options(clean_df, "Keterangan"),
            key="slik_filter_keterangan",
        )
        selected_bucket = c6.multiselect(
            "Maturity Bucket",
            [bucket for bucket in BUCKET_ORDER if bucket in set(clean_df["Maturity Bucket"].astype(str))],
            format_func=lambda value: BUCKET_LABELS.get(value, value),
            key="slik_filter_maturity",
        )

    mask = (
        _filter_by_values(clean_df, "Nama", selected_nama)
        & _filter_by_values(clean_df, "NIK/NPWP", selected_nik)
        & _filter_by_values(clean_df, "Bank", selected_bank)
        & _filter_by_values(clean_df, "Kol.", selected_kol)
        & _filter_by_values(clean_df, "Keterangan", selected_keterangan)
        & _filter_by_values(clean_df, "Maturity Bucket", selected_bucket)
    )
    filtered = clean_df.loc[mask].copy()

    if filtered.empty:
        st.warning("Tidak ada data yang cocok dengan filter.")
        return

    total_plafond = filtered["Plafond (Rp)"].sum()
    total_baki = filtered["Baki Debet (Rp)"].sum()
    active_mask = filtered["Keterangan"].fillna("").str.contains("aktif", case=False, na=False)
    writeoff_mask = filtered["Keterangan"].fillna("").str.contains(
        r"dihapusbukukan|hapus\s+tagih",
        case=False,
        na=False,
        regex=True,
    )
    problem_mask = filtered["Kol."].fillna(0).astype(int) >= 3
    due_soon_mask = filtered["Remaining Days"].between(0, 90).fillna(False)
    overdue_mask = (filtered["Remaining Days"] < 0).fillna(False)
    avg_rate = filtered["Rate"].dropna().mean()
    utilization = total_baki / total_plafond if total_plafond else 0
    worst_kol = filtered["Kol."].dropna().max()

    metrics = [
        ("Total Plafond", _format_rupiah(total_plafond)),
        ("Total Baki Debet", _format_rupiah(total_baki)),
        ("Jumlah Debitur", f"{filtered['Nama'].nunique():,}".replace(",", ".")),
        ("Jumlah Fasilitas", f"{len(filtered):,}".replace(",", ".")),
        ("Jumlah Bank / Lembaga Keuangan", f"{filtered['Bank'].nunique():,}".replace(",", ".")),
        ("Kol. Terburuk", "-" if pd.isna(worst_kol) else str(int(worst_kol))),
        ("Rata-rata Rate", _format_percent(avg_rate)),
        ("Rasio Utilisasi", _format_percent(utilization)),
        ("Fasilitas Aktif", f"{int(active_mask.sum()):,}".replace(",", ".")),
        ("Fasilitas Hapus Buku", f"{int(writeoff_mask.sum()):,}".replace(",", ".")),
        ("Exposure Kredit Bermasalah", _format_rupiah(filtered.loc[problem_mask, "Baki Debet (Rp)"].sum())),
        ("Fasilitas Jatuh Tempo 0-3 Bulan", f"{int(due_soon_mask.sum()):,}".replace(",", ".")),
        ("Fasilitas Lewat Jatuh Tempo", f"{int(overdue_mask.sum()):,}".replace(",", ".")),
    ]
    _metric_grid(metrics)

    st.subheader("Grafik")
    bank_baki = (
        filtered.groupby("Bank", dropna=False)["Baki Debet (Rp)"]
        .sum()
        .sort_values(ascending=False)
        .head(10)
        .reset_index()
    )
    bank_plafond = (
        filtered.groupby("Bank", dropna=False)["Plafond (Rp)"]
        .sum()
        .sort_values(ascending=False)
        .head(10)
        .reset_index()
    )
    kol_count = filtered.groupby("Kol.", dropna=False).size().reset_index(name="Jumlah Fasilitas")
    kol_baki = filtered.groupby("Kol.", dropna=False)["Baki Debet (Rp)"].sum().reset_index()
    debtor_baki = (
        filtered.groupby("Nama", dropna=False)["Baki Debet (Rp)"]
        .sum()
        .sort_values(ascending=False)
        .head(10)
        .reset_index()
    )
    maturity_summary = (
        filtered.groupby("Maturity Bucket", dropna=False)
        .agg(**{"Jumlah Fasilitas": ("Nama", "size"), "Baki Debet (Rp)": ("Baki Debet (Rp)", "sum")})
        .reset_index()
    )
    maturity_summary["Urutan"] = maturity_summary["Maturity Bucket"].map(
        {bucket: index for index, bucket in enumerate(BUCKET_ORDER)}
    ).fillna(999)
    maturity_summary = maturity_summary.sort_values("Urutan").drop(columns=["Urutan"])
    maturity_summary["Bucket"] = maturity_summary["Maturity Bucket"].map(BUCKET_LABELS).fillna(
        maturity_summary["Maturity Bucket"]
    )
    rate_df = filtered.loc[filtered["Rate"].notna(), ["Rate"]].copy()
    rate_df["Rate (%)"] = rate_df["Rate"] * 100

    c1, c2 = st.columns(2)
    with c1:
        _show_bar(bank_baki, "Bank", "Baki Debet (Rp)", "Total Baki Debet by Bank")
        _show_bar(kol_count, "Kol.", "Jumlah Fasilitas", "Facility Count by Kol.")
        _show_bar(debtor_baki, "Nama", "Baki Debet (Rp)", "Top 10 Debtors by Total Baki Debet")
        _show_bar(maturity_summary, "Bucket", "Jumlah Fasilitas", "Maturity Bucket summary")
    with c2:
        _show_bar(bank_plafond, "Bank", "Plafond (Rp)", "Total Plafond by Bank")
        _show_bar(kol_baki, "Kol.", "Baki Debet (Rp)", "Total Baki Debet by Kol.")
        _show_bar(bank_baki, "Bank", "Baki Debet (Rp)", "Top 10 Banks by Total Baki Debet")
        _show_histogram(rate_df, "Rate (%)", "Rate distribution")

    st.subheader("Detail Data Terfilter")
    st.dataframe(filtered, use_container_width=True, height=420)

    if errors_df is not None and not errors_df.empty:
        st.subheader("File Error")
        st.dataframe(errors_df, use_container_width=True)

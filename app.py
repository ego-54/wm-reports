import io
from datetime import datetime

import streamlit as st

import core

st.set_page_config(
    page_title="WineMore Reports",
    page_icon="🍷",
    layout="wide",
)

# ─── page style ──────────────────────────────────────────────────────────────

st.markdown("""
<style>
    .block-container { padding-top: 2rem; }
    .stTabs [data-baseweb="tab"] { font-size: 1rem; font-weight: 600; }
    .metric-box {
        background: #f8f9fa; border-radius: 8px;
        padding: 1rem 1.2rem; text-align: center;
    }
    .metric-box .val { font-size: 2rem; font-weight: 700; color: #1F4E79; }
    .metric-box .lbl { font-size: 0.8rem; color: #666; margin-top: 2px; }
</style>
""", unsafe_allow_html=True)

st.title("🍷 WineMore Reports")
st.caption("Internal tool — inventory reporting & import generation")

tab1, tab2 = st.tabs(["📊  Inventory Report", "📦  Inventory Import CSV"])


# ═══════════════════════════════════════════════════════════════════════════
# TAB 1 — INVENTORY REPORT
# ═══════════════════════════════════════════════════════════════════════════

with tab1:
    st.subheader("Inventory Report")
    st.markdown("Upload your Shopify exports and supplier files to generate the Excel report.")

    col1, col2 = st.columns(2)

    with col1:
        st.markdown("**Shopify Exports**")
        inv_files = st.file_uploader(
            "Inventory export(s)",
            type="csv",
            accept_multiple_files=True,
            key="inv_report",
            help="Shopify Admin → Inventory → Export",
        )
        prod_files = st.file_uploader(
            "Products export(s) — optional, adds Price",
            type="csv",
            accept_multiple_files=True,
            key="prod_report",
            help="Shopify Admin → Products → Export (can be 2 files)",
        )

    with col2:
        st.markdown("**Supplier Files**")
        joval_file = st.file_uploader(
            "Joval (CSV with Handle column)",
            type=["csv"],
            key="joval_report",
        )
        bib_file = st.file_uploader("Bibendum CSV", type="csv", key="bib_report")
        dws_file = st.file_uploader("DWS CSV",      type="csv", key="dws_report")
        sss_file = st.file_uploader("SSS CSV",      type="csv", key="sss_report")

    st.divider()

    if st.button("Generate Report", type="primary", key="gen_report"):
        if not inv_files:
            st.error("Please upload at least one inventory export CSV.")
        else:
            with st.spinner("Loading data and building report..."):

                # Load inventory
                inventory = core.load_inventory_from_files(inv_files)

                # Load products (optional)
                products = core.load_products_from_files(prod_files) if prod_files else {}

                # Build supplier map
                supplier_map = {}
                sup_sources = [
                    (joval_file, "joval"),
                    (bib_file,   "bibendum"),
                    (dws_file,   "dws"),
                    (sss_file,   "sss"),
                ]
                for f, label in sup_sources:
                    if f:
                        handles = core.load_supplier_handles_from_csv(f)
                        for h in handles:
                            if h in supplier_map:
                                supplier_map[h] += f", {label}"
                            else:
                                supplier_map[h] = label

                # Generate
                wb, stats = core.build_report(products, inventory, supplier_map)

                buf = io.BytesIO()
                wb.save(buf)
                buf.seek(0)

            # Metrics
            c1, c2, c3, c4, c5 = st.columns(5)
            c1.metric("Total Variants",  f"{stats['total']:,}")
            c2.metric("Stocked",         f"{stats['stocked']:,}")
            c3.metric("Zero Stock",      f"{stats['zero']:,}")
            c4.metric("WMC Units",       f"{stats['wmc_units']:,}")
            c5.metric("Chadstone Units", f"{stats['chad_units']:,}")

            dated = datetime.now().strftime("%Y-%m-%d")
            st.success("Report ready!")
            st.download_button(
                label="⬇️  Download Excel Report",
                data=buf,
                file_name=f"inventory_report_{dated}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )


# ═══════════════════════════════════════════════════════════════════════════
# TAB 2 — INVENTORY IMPORT CSV
# ═══════════════════════════════════════════════════════════════════════════

with tab2:
    st.subheader("Inventory Import CSV")
    st.markdown(
        "Upload supplier files with stock quantities and your inventory export. "
        "Downloads a Shopify-ready CSV you can import via **Shopify Admin → Inventory → Import**."
    )

    col1, col2 = st.columns(2)

    with col1:
        st.markdown("**Shopify Export**")
        inv_import_files = st.file_uploader(
            "Inventory export(s)",
            type="csv",
            accept_multiple_files=True,
            key="inv_import",
            help="Shopify Admin → Inventory → Export",
        )

    with col2:
        st.markdown("**Supplier Quantity Files**")
        joval_xlsx = st.file_uploader(
            "Joval XLSX (Qty. In Stock col G)",
            type=["xlsx"],
            key="joval_qty",
        )
        other_qty_files = st.file_uploader(
            "Other supplier qty CSVs (SKU + Quantity columns)",
            type="csv",
            accept_multiple_files=True,
            key="other_qty",
            help="e.g. bibendum_qty.csv, dws_qty.csv, sss_qty.csv",
        )

    st.divider()

    if st.button("Generate Import CSV", type="primary", key="gen_import"):
        if not inv_import_files:
            st.error("Please upload at least one inventory export CSV.")
        else:
            all_qty = {}

            if joval_xlsx:
                qty = core.load_joval_quantities_xlsx(joval_xlsx)
                all_qty.update(qty)
                st.info(f"Joval: {len(qty):,} SKUs loaded")

            for f in (other_qty_files or []):
                qty, err = core.load_qty_csv(f)
                if err:
                    st.warning(f"{f.name}: {err}")
                else:
                    all_qty.update(qty)
                    st.info(f"{f.name}: {len(qty):,} SKUs loaded")

            if not all_qty:
                st.error("No quantities found. Upload joval.xlsx or a qty CSV with SKU + Quantity columns.")
            else:
                with st.spinner("Building import CSV..."):
                    inv_rows = core.load_flat_inventory_rows(inv_import_files)
                    csv_bytes, stats, unmatched = core.build_inventory_import(inv_rows, all_qty)

                c1, c2, c3 = st.columns(3)
                c1.metric("Variants Updated",      f"{stats['matched']:,}")
                c2.metric("Skipped (no match)",    f"{stats['skipped']:,}")
                c3.metric("Supplier SKUs not found", f"{stats['not_found']:,}")

                dated = datetime.now().strftime("%Y-%m-%d_%H%M")
                st.success("Import CSV ready!")
                st.download_button(
                    label="⬇️  Download Shopify Import CSV",
                    data=csv_bytes,
                    file_name=f"shopify_inventory_import_{dated}.csv",
                    mime="text/csv",
                )

                if unmatched:
                    with st.expander(f"⚠️  {len(unmatched):,} supplier SKUs not found in Shopify"):
                        st.code("\n".join(unmatched[:100]))
                        if len(unmatched) > 100:
                            st.caption(f"Showing first 100 of {len(unmatched)}")

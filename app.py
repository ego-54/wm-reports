import io
from datetime import datetime

import streamlit as st

import core

st.set_page_config(
    page_title="WineMore Reports",
    page_icon="🍷",
    layout="wide",
)

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
            type=["csv", "xlsx"],
            accept_multiple_files=True,
            key="inv_report",
            help="Shopify Admin → Inventory → Export",
        )
        prod_files = st.file_uploader(
            "Products export(s) — optional, adds Price",
            type=["csv", "xlsx"],
            accept_multiple_files=True,
            key="prod_report",
            help="Shopify Admin → Products → Export (can be 2 files)",
        )

    with col2:
        st.markdown("**Supplier Files**")
        joval_file = st.file_uploader(
            "Joval (CSV with Handle column)",
            type=["csv", "xlsx"],
            key="joval_report",
        )
        bib_file = st.file_uploader("Bibendum CSV", type=["csv", "xlsx"], key="bib_report")
        dws_file = st.file_uploader("DWS CSV",      type=["csv", "xlsx"], key="dws_report")
        sss_file = st.file_uploader("SSS CSV",      type=["csv", "xlsx"], key="sss_report")

    st.divider()

    if st.button("Generate Report", type="primary", key="gen_report"):
        if not inv_files:
            st.error("Please upload at least one inventory export CSV.")
        else:
            with st.spinner("Loading data and building report..."):
                inventory = core.load_inventory_from_files(inv_files)
                products  = core.load_products_from_files(prod_files) if prod_files else {}

                supplier_map = {}
                for f, label in [(joval_file, "joval"), (bib_file, "bibendum"),
                                  (dws_file, "dws"), (sss_file, "sss")]:
                    if f:
                        handles = core.load_supplier_handles_from_csv(f)
                        for h in handles:
                            supplier_map[h] = (supplier_map[h] + f", {label}"
                                               if h in supplier_map else label)

                wb, stats = core.build_report(products, inventory, supplier_map)
                buf = io.BytesIO()
                wb.save(buf)
                buf.seek(0)

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
        "Upload supplier files with stock quantities and your Shopify exports. "
        "Downloads a Shopify-ready CSV you can import via **Admin → Inventory → Import**."
    )

    with st.expander("ℹ️  Joval quantity rules", expanded=False):
        st.markdown("""
        **Quantity** = `floor(Qty in stock × carton size)`

        **Zero thresholds by price tier:**
        | Price | Min units to list |
        |-------|------------------|
        | Under $50 | 30 |
        | $50 – $200 | 10 |
        | Over $200 | 5 |

        Any `supplier-joval` SKU **missing** from the SOH file is set to **0**.
        """)

    col1, col2 = st.columns(2)

    with col1:
        st.markdown("**Shopify Exports**")
        prod_import_files = st.file_uploader(
            "Products export(s)",
            type=["csv", "xlsx"],
            accept_multiple_files=True,
            key="prod_import",
            help="Shopify Admin → Products → Export (can be 2 files). Used for variant structure + Joval price rules.",
        )

    with col2:
        st.markdown("**Supplier Quantity Files**")
        joval_xlsx = st.file_uploader(
            "Joval XLSX (SOH file)",
            type=["csv", "xlsx"],
            key="joval_qty",
            help="Headers row 8 — Item No. (col A), Carton Size (col D), Qty In Stock (col G)",
        )
        other_qty_files = st.file_uploader(
            "Other supplier qty CSVs (SKU + Quantity columns)",
            type=["csv", "xlsx"],
            accept_multiple_files=True,
            key="other_qty",
        )

    st.divider()

    if st.button("Generate Import CSV", type="primary", key="gen_import"):
        if not prod_import_files:
            st.error("Please upload at least one products export CSV/XLSX.")
        else:
            all_qty    = {}
            sku_prices = {}
            joval_skus = set()

            sku_prices = core.load_sku_prices(prod_import_files)
            sku_titles = core.load_sku_titles(prod_import_files)
            joval_skus = core.load_supplier_skus(prod_import_files, tag="supplier-joval")
            st.info(f"Products: {len(sku_prices):,} SKU prices loaded, "
                    f"{len(joval_skus):,} supplier-joval SKUs found")

            joval_stats = None
            if joval_xlsx:
                qty, joval_stats = core.load_joval_quantities_xlsx(
                    joval_xlsx, sku_prices=sku_prices, sku_titles=sku_titles
                )
                all_qty.update(qty)
                st.info(f"Joval SOH: {joval_stats['total']:,} SKUs — "
                        f"{joval_stats['nonzero']:,} with stock, "
                        f"{len(joval_stats['zeroed_by_rules']):,} zeroed by price rules, "
                        f"{len(joval_stats['zeroed_raw']):,} zero raw stock")

            for f in (other_qty_files or []):
                qty, err = core.load_qty_csv(f)
                if err:
                    st.warning(f"{f.name}: {err}")
                else:
                    all_qty.update(qty)
                    st.info(f"{f.name}: {len(qty):,} SKUs loaded")

            if not all_qty and not joval_skus:
                st.error("No quantities found. Upload Joval XLSX or a qty CSV.")
            else:
                with st.spinner("Building import CSV..."):
                    inv_rows   = core.load_flat_rows_from_products(prod_import_files)
                    csv_bytes, stats, unmatched = core.build_inventory_import(
                        inv_rows, all_qty, joval_skus=joval_skus
                    )

                st.success("Import CSV ready!")

                # ── Overview metrics ──────────────────────────────────────────
                c1, c2, c3, c4, c5 = st.columns(5)
                c1.metric("Total in CSV",          f"{stats['matched']:,}")
                c2.metric("With Stock",             f"{stats['nonzero']:,}")
                c3.metric("Set to Zero",            f"{stats['zeroed']:,}")
                c4.metric("Not in Products Export", f"{stats['not_found']:,}")
                c5.metric("Supplier SKUs Missing",  f"{len(stats['zeroed_missing_skus']):,}")

                dated = datetime.now().strftime("%Y-%m-%d")
                st.download_button(
                    label="⬇️  Download inventory_joval CSV",
                    data=csv_bytes,
                    file_name=f"inventory_joval_{dated}.csv",
                    mime="text/csv",
                )

                st.divider()

                # ── Price tier breakdown ──────────────────────────────────────
                if joval_stats and joval_stats["zeroed_by_rules"]:
                    rules = joval_stats["zeroed_by_rules"]
                    tier_summary = {}
                    thresholds = {"Under $50": 30, "$50–$200": 10, "Over $200": 5}
                    for row in rules:
                        t = row["Tier"]
                        tier_summary[t] = tier_summary.get(t, 0) + 1
                    tier_rows = [
                        {"Price Tier": t, "Min Units Required": thresholds.get(t, "—"),
                         "SKUs Zeroed": tier_summary.get(t, 0)}
                        for t in ["Under $50", "$50–$200", "Over $200"] if t in tier_summary
                    ]
                    with st.expander(f"📊  Price rule zeroes — {len(rules):,} SKUs", expanded=True):
                        st.dataframe(tier_rows, hide_index=True, use_container_width=True)
                        st.caption("Full list:")
                        st.dataframe(rules, hide_index=True, use_container_width=True)

                # ── Zeroed because missing from SOH ───────────────────────────
                if stats["zeroed_missing_skus"]:
                    miss = stats["zeroed_missing_skus"]
                    with st.expander(f"⚠️  Zeroed — missing from SOH — {len(miss):,} SKUs"):
                        st.dataframe(
                            [{"SKU": s, "Product": sku_titles.get(s, "")} for s in miss],
                            hide_index=True, use_container_width=True,
                        )

                # ── Zero raw stock in SOH ─────────────────────────────────────
                if joval_stats and joval_stats["zeroed_raw"]:
                    raw = joval_stats["zeroed_raw"]
                    with st.expander(f"ℹ️  Zero stock in SOH — {len(raw):,} SKUs"):
                        st.dataframe(
                            [{"SKU": s, "Product": sku_titles.get(s, "")} for s in raw],
                            hide_index=True, use_container_width=True,
                        )

                # ── Preview table ─────────────────────────────────────────────
                if stats.get("preview"):
                    with st.expander("👁️  Preview — first 30 rows", expanded=False):
                        st.dataframe(stats["preview"], hide_index=True, use_container_width=True)

                # ── Supplier SKUs not matched in product export ───────────────
                if unmatched:
                    with st.expander(f"🔍  {len(unmatched):,} supplier SKUs not found in product export"):
                        st.dataframe([{"SKU": s} for s in unmatched[:200]],
                                     hide_index=True, use_container_width=True)
                        if len(unmatched) > 200:
                            st.caption(f"Showing first 200 of {len(unmatched)}")

# =====================================================
# FILTERS (WITH SELECT ALL / CLEAR ALL)
# =====================================================
with st.sidebar:
    st.header("🎯 Filter Data")

    # ---------- KATEGORI ----------
    st.subheader("Kategori")

    if "cat_select" not in st.session_state:
        st.session_state.cat_select = list(cont_cat.keys())

    cat_mode = st.radio(
        "Kategori Mode",
        ("Select All", "Clear All"),
        key="cat_mode"
    )

    if cat_mode == "Select All":
        st.session_state.cat_select = list(cont_cat.keys())
    else:
        st.session_state.cat_select = []

    st.session_state.cat_select = st.multiselect(
        "Pilih Kategori",
        options=list(cont_cat.keys()),
        default=st.session_state.cat_select
    )

    st.divider()

    # ---------- FORMAT ----------
    st.subheader("Format")

    if "fmt_select" not in st.session_state:
        st.session_state.fmt_select = list(cont_fmt.keys())

    fmt_mode = st.radio(
        "Format Mode",
        ("Select All", "Clear All"),
        key="fmt_mode"
    )

    if fmt_mode == "Select All":
        st.session_state.fmt_select = list(cont_fmt.keys())
    else:
        st.session_state.fmt_select = []

    st.session_state.fmt_select = st.multiselect(
        "Pilih Format",
        options=list(cont_fmt.keys()),
        default=st.session_state.fmt_select
    )

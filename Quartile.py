with st.sidebar:
    st.header("🎯 Filter Data")

    # =====================
    # KATEGORI
    # =====================
    st.subheader("Kategori")

    if "cat_select" not in st.session_state:
        st.session_state.cat_select = list(cont_cat.keys())

    col1, col2 = st.columns(2)
    with col1:
        if st.button("Select All", key="cat_all"):
            st.session_state.cat_select = list(cont_cat.keys())
    with col2:
        if st.button("Clear All", key="cat_clear"):
            st.session_state.cat_select = []

    cat_select = st.multiselect(
        "Pilih Kategori",
        options=list(cont_cat.keys()),
        default=st.session_state.cat_select,
        key="cat_select"
    )

    # =====================
    # FORMAT
    # =====================
    st.subheader("Format")

    if "fmt_select" not in st.session_state:
        st.session_state.fmt_select = list(cont_fmt.keys())

    col3, col4 = st.columns(2)
    with col3:
        if st.button("Select All", key="fmt_all"):
            st.session_state.fmt_select = list(cont_fmt.keys())
    with col4:
        if st.button("Clear All", key="fmt_clear"):
            st.session_state.fmt_select = []

    fmt_select = st.multiselect(
        "Pilih Format",
        options=list(cont_fmt.keys()),
        default=st.session_state.fmt_select,
        key="fmt_select"
    )

import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="S&OP Editor", layout="wide")

uploaded_file = st.file_uploader("Carga el archivo S&OP", type=["xlsx"])

if uploaded_file:
    sheets = pd.read_excel(uploaded_file, sheet_name=None)
    hojas_editables = {name: df.copy() for name, df in sheets.items()}

    hoja = st.selectbox("Selecciona hoja", list(hojas_editables.keys()))
    df = hojas_editables[hoja]

    col1, col2, col3, col4, col5 = st.columns(5)
    depto = col1.selectbox("Department", ["Todos"] + sorted(df['DEPARTMENT'].dropna().unique().tolist()))
    category = col2.selectbox("Category", ["Todos"] + sorted(df['CATEGORY'].dropna().unique().tolist()))
    supplier = col3.selectbox("Supplier", ["Todos"] + sorted(df['SUPPLIER'].dropna().unique().tolist()))
    product = col4.selectbox("Product", ["Todos"] + sorted(df['PRODUCT'].dropna().unique().tolist()))

    if hoja.lower() == "directo" and "TIENDA" in df.columns:
        location_field = "TIENDA"
    elif hoja.lower() == "centralizado" and "CEDIS Entrega" in df.columns:
        location_field = "CEDIS Entrega"
    else:
        location_field = None

    if location_field:
        location_values = ["Todos"] + sorted(df[location_field].dropna().unique().tolist())
        location = col5.selectbox(location_field, location_values)
    else:
        location = "Todos"

    df_filtro = df.copy()
    if depto != "Todos": df_filtro = df_filtro[df_filtro['DEPARTMENT'] == depto]
    if category != "Todos": df_filtro = df_filtro[df_filtro['CATEGORY'] == category]
    if supplier != "Todos": df_filtro = df_filtro[df_filtro['SUPPLIER'] == supplier]
    if product != "Todos": df_filtro = df_filtro[df_filtro['PRODUCT'] == product]
    if location_field and location != "Todos":
        df_filtro = df_filtro[df_filtro[location_field] == location]

    st.markdown("### Edita la columna DOH_TARGET")
    editable_cols = ['DEPARTMENT', 'CATEGORY', 'SUPPLIER', 'PRODUCT']

    # Insertar columnas específicas según hoja
    if hoja.lower() == "centralizado":
        central_cols = ['INV + TRANSIT', 'CEDIS_ORDERED_UNITS', 'INV ALMACEN', 'TTL INV']
        editable_cols.extend(central_cols)
    elif hoja.lower() == "directo":
        directo_cols = ['INV TIENDA', 'TRANSITO', 'INV + TRANSIT']
        editable_cols.extend(directo_cols)

    editable_cols.extend(['DOH_TARGET', 'VENTA REAL PROM'])
    if location_field and location_field not in editable_cols:
        editable_cols.insert(0, location_field)

    df_edit = st.data_editor(df_filtro[editable_cols], num_rows="dynamic", use_container_width=True, key="editor")

    if st.button("Recalcular y actualizar"):
        for _, row in df_edit.iterrows():
            cond = (
                (df['DEPARTMENT'] == row['DEPARTMENT']) &
                (df['CATEGORY'] == row['CATEGORY']) &
                (df['SUPPLIER'] == row['SUPPLIER']) &
                (df['PRODUCT'] == row['PRODUCT'])
            )
            if location_field in row:
                cond &= df[location_field] == row[location_field]
            df.loc[cond, 'DOH_TARGET'] = row['DOH_TARGET']

        df['INV TARGET'] = df['DOH_TARGET'] * df['VENTA REAL PROM']
        df['COMPRA'] = df['INV TARGET'] - df['TTL INV']
        df['COMPRA'] = df['COMPRA'].apply(lambda x: max(0, x))

        def compra_umi(row):
            if row['UMI'] == 'KG':
                return (row['COMPRA'] / row['QUANTITY_PER_UMI']) * row['QUANTITY_PER_UMI']
            else:
                return round(row['COMPRA'] / row['QUANTITY_PER_UMI'])

        df['COMPRA UMI'] = df.apply(compra_umi, axis=1)
        df['DOH COMPRA'] = df['COMPRA'] / df['VENTA REAL PROM']
        df['DOH ACTUAL'] = (df['TTL INV'] + df['COMPRA']) / df['VENTA REAL PROM']
        df['DOH FINALES'] = df['DOH COMPRA'] + df['DOH ACTUAL']

        st.success("Valores recalculados exitosamente.")

        df_filtro_actualizado = df.copy()
        if depto != "Todos": df_filtro_actualizado = df_filtro_actualizado[df_filtro_actualizado['DEPARTMENT'] == depto]
        if category != "Todos": df_filtro_actualizado = df_filtro_actualizado[df_filtro_actualizado['CATEGORY'] == category]
        if supplier != "Todos": df_filtro_actualizado = df_filtro_actualizado[df_filtro_actualizado['SUPPLIER'] == supplier]
        if product != "Todos": df_filtro_actualizado = df_filtro_actualizado[df_filtro_actualizado['PRODUCT'] == product]
        if location_field and location != "Todos":
            df_filtro_actualizado = df_filtro_actualizado[df_filtro_actualizado[location_field] == location]

        mostrar_cols = [location_field, 'DEPARTMENT', 'CATEGORY', 'SUPPLIER', 'PRODUCT'] if location_field else [
                        'DEPARTMENT', 'CATEGORY', 'SUPPLIER', 'PRODUCT']

        if hoja.lower() == "centralizado":
            mostrar_cols += ['INV + TRANSIT', 'CEDIS_ORDERED_UNITS', 'INV ALMACEN', 'TTL INV']
        elif hoja.lower() == "directo":
            mostrar_cols += ['INV TIENDA', 'TRANSITO', 'INV + TRANSIT']

        mostrar_cols += ['DOH_TARGET', 'VENTA REAL PROM', 'COMPRA', 'COMPRA UMI', 'DOH ACTUAL', 'DOH COMPRA', 'DOH FINALES']

        st.dataframe(df_filtro_actualizado[mostrar_cols], use_container_width=True)

        def to_excel(dataframe):
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                dataframe.to_excel(writer, index=False, sheet_name='Actualizado')
            return output.getvalue()

        excel_data = to_excel(df)
        st.download_button(
            label="📥 Descargar archivo actualizado",
            data=excel_data,
            file_name="SOP_actualizado.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

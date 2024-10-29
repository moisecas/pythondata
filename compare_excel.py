import pandas as pd

# Definir las rutas de los archivos de Excel
file1 = 'C:\\Users\\moise\\Downloads\\reporteutilidadesmigra.xlsx'
file2 = 'C:\\Users\\moise\\Downloads\\reporteutilidadesprod.xlsx'

# Leer los archivos de Excel
df1 = pd.read_excel(file1, sheet_name=None)
df2 = pd.read_excel(file2, sheet_name=None)

# Verificar que las hojas tengan el mismo nombre
sheets1 = df1.keys()
sheets2 = df2.keys()

if sheets1 != sheets2:
    print("Los archivos tienen diferentes hojas.")
else:
    # Comparar las hojas
    for sheet in sheets1:
        data1 = df1[sheet]
        data2 = df2[sheet]
        
        # Alinear los DataFrames por índice y columnas
        data1, data2 = data1.align(data2, join='outer', axis=1)  # Alineación por columnas
        data1, data2 = data1.align(data2, join='outer', axis=0)  # Alineación por filas

        # Llenar valores NaN con un valor neutral (opcional)
        data1 = data1.fillna('NaN')
        data2 = data2.fillna('NaN')

        # Comparar los datos de las hojas
        comparison = data1.compare(data2, keep_equal=False)

        if not comparison.empty:  # Verifica si hay diferencias
            # Guardar el reporte de diferencias si existen
            report_file = f'reporte_diferencias_{sheet}.xlsx'
            comparison.to_excel(report_file)
            print(f"Diferencias encontradas en la hoja '{sheet}'. Reporte guardado en {report_file}")
        else:
            print(f"No se encontraron diferencias en la hoja '{sheet}'.")


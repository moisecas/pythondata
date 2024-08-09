import pandas as pd

# Definir las rutas de los archivos de Excel
file1 = 'C:\\Users\\moise\\Downloads\\informe-vacacion-periodo-peru_despues.xls'
file2 = 'C:\\Users\\moise\\Downloads\\informe-vacacion-periodo-peru_antes.xls'


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

        # Comparar los datos de las hojas
        comparison = data1.compare(data2, keep_equal=False)

        # Guardar el reporte de diferencias
        report_file = f'reporte_diferencias_{sheet}.xlsx'
        comparison.to_excel(report_file)
        print(f'Reporte de diferencias guardado en {report_file}')

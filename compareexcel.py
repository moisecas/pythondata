import pandas as pd

# Leer los archivos de Excel cuando tienen columnas o celdas diferentes
file1 = 'C:\\Users\\moise\\Downloads\\informe-vacacion-periodo-peru_despues.xls'
file2 = 'C:\\Users\\moise\\Downloads\\informe-vacacion-periodo-peru_antes.xls'

# Cargar los datos de las hojas
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
        
        # Alinear los DataFrames para asegurarse de que tengan las mismas columnas e índices
        data1, data2 = data1.align(data2, join='outer', axis=1)
        data1, data2 = data1.align(data2, join='outer', axis=0)
        
        # Comparar los datos de las hojas
        comparison = data1.compare(data2, keep_equal=False)

        # Guardar el reporte de diferencias
        report_file = f'reporte_diferencias_{sheet}.xlsx'
        comparison.to_excel(report_file)
        print(f'Reporte de diferencias guardado en {report_file}')

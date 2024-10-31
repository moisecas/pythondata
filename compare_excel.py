import pandas as pd

# Definir las rutas de los archivos de Excel
migra = 'C:\\Users\\moise\\Downloads\\reporte-provision-fijo-gratificacion-12-2023 (3).xls'
prod = 'C:\\Users\\moise\\Downloads\\reporte-provision-fijo-gratificacion-12-2023 (2).xls'

# Leer los archivos de Excel
migra_data = pd.read_excel(migra, sheet_name=None)
prod_data = pd.read_excel(prod, sheet_name=None)

# Verificar que las hojas tengan el mismo nombre 
migra_sheets = migra_data.keys()
prod_sheets = prod_data.keys()

if migra_sheets != prod_sheets:
    print("Los archivos tienen diferentes hojas.")
else:
    # Comparar las hojas
    for sheet in migra_sheets:
        migra_df = migra_data[sheet]
        prod_df = prod_data[sheet]
        
        # Alinear los DataFrames por índice y columnas
        migra_df, prod_df = migra_df.align(prod_df, join='outer', axis=1)  # Alineación por columnas
        migra_df, prod_df = migra_df.align(prod_df, join='outer', axis=0)  # Alineación por filas

        # Llenar valores NaN con un valor neutral (opcional)
        migra_df = migra_df.fillna('NaN')
        prod_df = prod_df.fillna('NaN')

        # Comparar los datos de las hojas y agregar detalles sobre las diferencias
        comparison = migra_df.compare(prod_df, keep_equal=False, result_names=('Migra', 'Prod'))

        if not comparison.empty:  # Verifica si hay diferencias
            # Guardar el reporte de diferencias si existen, incluyendo valores previos y nuevos
            report_file = f'reporte_diferencias_{sheet}.xlsx'
            with pd.ExcelWriter(report_file) as writer:
                comparison.to_excel(writer, sheet_name='Diferencias')
                migra_df.to_excel(writer, sheet_name='Migra')
                prod_df.to_excel(writer, sheet_name='Prod')
                
            print(f"Diferencias encontradas en la hoja '{sheet}'. Reporte detallado guardado en {report_file}")
        else:
            print(f"No se encontraron diferencias en la hoja '{sheet}'.")


import openpyxl
from openpyxl import Workbook

# Función para obtener los detalles del gasto del usuario
def ingresar_gasto():
    fecha = input("Ingresa la fecha del gasto (dd/mm/aaaa): ")
    descripcion = input("Ingresa la descripción del gasto: ")
    monto = float(input("Ingresa el monto del gasto: "))
    return {"fecha": fecha, "descripcion": descripcion, "monto": monto}

# Crear o cargar el archivo Excel
def crear_archivo_excel(nombre_archivo):
    try:
        workbook = openpyxl.load_workbook(nombre_archivo)
        print(f"Archivo {nombre_archivo} cargado correctamente.")
    except FileNotFoundError:
        workbook = Workbook()
        print(f"Archivo {nombre_archivo} creado.")
    return workbook

# Guardar datos en la hoja
def guardar_gastos(workbook, gastos, nombre_archivo):
    hoja = workbook.active
    hoja.title = "Gastos"
    hoja.append(["Fecha", "Descripción", "Monto"])
    
    for gasto in gastos:
        hoja.append([gasto["fecha"], gasto["descripcion"], gasto["monto"]])
    
    workbook.save(nombre_archivo)
    print(f"Datos guardados en {nombre_archivo}")

# Generar resumen de gastos
def generar_resumen(gastos):
    if not gastos:
        print("No se ingresaron gastos.")
        return
    
    total_gastos = sum(gasto["monto"] for gasto in gastos)
    gasto_mas_caro = max(gastos, key=lambda x: x["monto"])
    gasto_mas_barato = min(gastos, key=lambda x: x["monto"])
    
    print("\n--- Resumen de Gastos ---")
    print(f"Número total de gastos: {len(gastos)}")
    print(f"Gasto más caro: {gasto_mas_caro['fecha']} - {gasto_mas_caro['descripcion']}: {gasto_mas_caro['monto']}")
    print(f"Gasto más barato: {gasto_mas_barato['fecha']} - {gasto_mas_barato['descripcion']}: {gasto_mas_barato['monto']}")
    print(f"Monto total de gastos: {total_gastos}")
    
# Programa principal
def main():
    nombre_archivo = "informe_gastos.xlsx"
    workbook = crear_archivo_excel(nombre_archivo)
    
    gastos = []
    
    while True:
        gastos.append(ingresar_gasto())
        continuar = input("¿Deseas agregar otro gasto? (s/n): ").lower()
        if continuar != 's':
            break
    
    guardar_gastos(workbook, gastos, nombre_archivo)
    generar_resumen(gastos)

# Ejecutar el programa
if __name__ == "__main__":
    main()

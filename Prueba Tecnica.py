import sqlite3
import win32com.client as win32
from datetime import datetime
from openpyxl import Workbook
import os

conn = sqlite3.connect('dataBase/database.sqlite')
cursor = conn.cursor()

query = """
SELECT 
    c.commerce_id,
    c.commerce_name,
    SUM(CASE WHEN a.ask_status = 'Successful' THEN 1 ELSE 0 END) AS total_exitosas,
    SUM(CASE WHEN a.ask_status = 'Unsuccessful' THEN 1 ELSE 0 END) AS total_no_exitosas
FROM apicall a
JOIN commerce c ON a.commerce_id = c.commerce_id
WHERE c.commerce_status = 'Active' AND strftime('%Y-%m', a.date_api_call) IN ('2024-07', '2024-08')
GROUP BY c.commerce_id, c.commerce_name;
"""
cursor.execute(query)
data = cursor.fetchall()

query_emails = """
SELECT 
    commerce_id,
    commerce_email,
    commerce_nit
FROM commerce
WHERE commerce_status = 'Active';
"""
cursor.execute(query_emails)
email_data = cursor.fetchall()
email_dicc = {row[0]: {"email": row[1], "nit": row[2]} for row in email_data}

terms_contrato = {
    "Innovexa Solutions": {
        "tipo": "fijo",
        "precio": 300,
        "iva": 0.19
    },
    "NexaTech Industries": {
        "tipo": "variable",
        "categoria": [
            {"max_requests": 10000, "precio": 250},
            {"max_requests": 20000, "precio": 200},
            {"max_requests": float("inf"), "precio": 170}
        ],
        "iva": 0.19
    },
    "QuantumLeap Inc.": {
        "tipo": "fijo",
        "precio": 600,
        "iva": 0.19
    },
    "Zenith Corp.": {
        "tipo": "variable",
        "categoria": [
            {"max_requests": 22000, "precio": 250},
            {"max_requests": float("inf"), "precio": 130}
        ],
        "iva": 0.19,
        "descuentos": [
            {"condition": lambda no_exitosas: no_exitosas > 6000, "descuento": 0.05}
        ]
    },
    "FusionWave Enterprises": {
        "tipo": "fijo",
        "precio": 300,
        "iva": 0.19,
        "descuentos": [
            {"condition": lambda no_exitosas: 2500 <= no_exitosas <= 4500, "descuento": 0.05},
            {"condition": lambda no_exitosas: no_exitosas > 4500, "descuento": 0.08}
        ]
    }
}

def calcular_cobro(comercio, exitosas, no_exitosas):
    if comercio not in terms_contrato:
        return 0 
    
    contrato = terms_contrato[comercio]
    iva = contrato["iva"]
    total_bruto = 0
    desc = 0

    if contrato["tipo"] == "fijo":
        total_bruto = exitosas * contrato["precio"]
    
    elif contrato["tipo"] == "variable":
        for tier in contrato["categoria"]:
            if exitosas <= tier["max_requests"]:
                total_bruto = exitosas * tier["precio"]
                break

    # Aplicar descuentos si existen
    if "descuentos" in contrato:
        for descuento in contrato["descuentos"]:
            if descuento["condition"](no_exitosas):
                desc = max(desc, descuento["descuento"])

    total_descuento = total_bruto * desc
    total_neto = (total_bruto - total_descuento) * (1 + iva)
    
    return round(total_neto, 2)

# Depurador de los datos obtenidos de la consulta SQL
resultados_cobro = []
# for comercio_id, comercio, exitosas, no_exitosas in data:
#     print(f"Depuración: Comercio: {comercio}, Exitosas: {exitosas}, No exitosas: {no_exitosas}")
#     total_cobrar = calcular_cobro(comercio, exitosas, no_exitosas)
#     resultados_cobro.append((comercio, total_cobrar))
# Mostrar resultados
for comercio, total in resultados_cobro:
    print(f"Comercio: {comercio} - Total a cobrar: ${total} COP")
    
    resultados_cobro = []
for comercio_id, comercio, exitosas, no_exitosas in data:
    if comercio_id in email_dicc: 
        total_cobrar = calcular_cobro(comercio, exitosas, no_exitosas)
        valor_comision = total_cobrar / 0.19  # Valor sin IVA
        valor_iva = total_cobrar - valor_comision
        resultados_cobro.append({
            "Fecha-Mes": "Julio-Agosto 2024",
            "Nombre": comercio,
            "Nit": email_dicc[comercio_id]["nit"],
            "Valor_comision": round(valor_comision, 2),
            "Valor_iva": round(valor_iva, 2),
            "Valor_Total": round(total_cobrar, 2),
            "Correo": email_dicc[comercio_id]["email"]
        })
        
        # Ruta absoluta del archivo Excel
output_file = os.path.abspath("resultado/facturas_resumen.xlsx")

# Inicializar el archivo Excel
wb = Workbook()
ws = wb.active
ws.title = "Facturas Resumen"

#Encabezados
headers = ["Fecha-Mes", "Nombre", "Nit", "Valor Comisión", "Valor IVA", "Valor Total", "Correo"]
ws.append(headers)

for resultado in resultados_cobro:
    ws.append([
        resultado["Fecha-Mes"],
        resultado["Nombre"],
        resultado["Nit"],
        resultado["Valor_comision"],
        resultado["Valor_iva"],
        resultado["Valor_Total"],
        resultado["Correo"]
    ])

# Guardar archivo Excel :D
wb.save(output_file)
print(f"Archivo de excel guardado en: {output_file}")

# Crear tabla HTML
tabla_html = """
<table border="1" style="border-collapse: collapse; text-align: left;">
    <tr>
        <th>Fecha-Mes</th>
        <th>Nombre</th>
        <th>NIT</th>
        <th>Valor Comisión</th>
        <th>Valor IVA</th>
        <th>Valor Total</th>
        <th>Correo</th>
    </tr>
"""
for resultado in resultados_cobro:
    tabla_html += f"""
    <tr>
        <td>{resultado['Fecha-Mes']}</td>
        <td>{resultado['Nombre']}</td>
        <td>{resultado['Nit']}</td>
        <td>${resultado['Valor_comision']}</td>
        <td>${resultado['Valor_iva']}</td>
        <td>${resultado['Valor_Total']}</td>
        <td>{resultado['Correo']}</td>
    </tr>
    """
tabla_html += "</table>"

# Lanzar Outlook y definir el correo
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = "valentina.meneses033@gmail.com" 
mail.Subject = "Resumen de Facturas - Julio y Agosto 2024"
mail.HTMLBody = f"""
<p>Holitas!!,</p>
<p>A continuación, se presenta el resumen de las facturas generadas para los meses de Julio y Agosto de 2024 :) </p>
{tabla_html}
<p>Saludos!,</p>
"""

# Adjuntamos y enviamos el archivo Excel
mail.Attachments.Add(output_file)
mail.Send()

print("Correo enviado exitosamente.")
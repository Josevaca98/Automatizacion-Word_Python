from docxtpl import DocxTemplate
from datetime import datetime
import pandas as pd
import os


pathPrincipal = "./Proyecto_02_AutomatizarWord/"
pathSalida = os.path.join(pathPrincipal, "documentos_generados")

os.makedirs(pathSalida,exist_ok=True)

doc = DocxTemplate(os.path.join(pathPrincipal, "sp-plantilla-rrhh-info.docx"))

mi_nombre = "Jose María Vaca González"
mi_numero = "(444) 433-75-40"
mi_correo = "licjmvg98@gmail.com"
mi_direccion = "Los Vargas #127, San Luis potosí, S.L.P"
fecha_hoy = datetime.today().strftime("%d %b, %Y")


my_context = { 'mi_nombre' : mi_nombre,
           'mi_numero': mi_numero,
           'mi_correo': mi_correo,
           'mi_direccion': mi_direccion,
           'fecha_hoy': fecha_hoy
           }

df = pd.read_csv(os.path.join(pathPrincipal, "sp_fake_data.csv"))

for index, fila in df.iterrows():
    context = {
        'nombre_persona_rrhh' : fila['name'],
        'direccion' : fila['address'],
        'numero_telefono' : fila['phone_number'],
        'correo' : fila['email'],
        'puesto_trabajo': fila['job'] ,
        'nombre_empresa' : fila['company'], 
    }
    context.update(my_context)

    doc.render(context)
    output_path = os.path.join(pathSalida, f"doc_generado_{index}.docx")
    doc.save(output_path)

print(f"Documentos generados en: {pathSalida}")
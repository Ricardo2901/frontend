
from capitulo17 import capitulo17

def crear_capitulos_proyecto(nombre_proyecto):
    print(f"📂 Generando proyecto: {nombre_proyecto}...")
    capitulo17(nombre_proyecto)

    print(f"✅ Proyecto '{nombre_proyecto}' generado con todos los capítulos.")

if __name__ == "__main__":
    nombre_proyecto = input("\nQue tipo de proyecto deseas generar?")
    crear_capitulos_proyecto(nombre_proyecto)

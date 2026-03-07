import pandas as pd
import random

datos = []
num_costaleros = 60
opciones = ["Izquierdo", "Derecho", "Ambos"]
pesos = [15, 15, 70] 

for i in range(1, num_costaleros + 1):
    nombre = f"Costalero {i}"
    preferencia = random.choices(opciones, weights=pesos, k=1)[0]
    
    # Altura base (hombro izquierdo)
    # Rango: desde 120cm (pers. 1.50m) hasta 175cm (pers. 2.00m)
    h_izq = round(random.uniform(120.0, 175.0), 2)
    
    # Simulación médica real:
    # La mayoría (80%) varía entre 0.5 y 1.2 cm. 
    # Un pequeño grupo (20%) varía hasta 2.5 cm (posturas, escoliosis leve).
    if random.random() < 0.8:
        variacion = random.uniform(0.5, 1.2)
    else:
        variacion = random.uniform(1.2, 2.5)
    
    # Aplicamos la regla del hombro dominante: suele estar más bajo
    # Si es diestro, el derecho baja. Si es zurdo, el izquierdo baja.
    if preferencia == "Derecho":
        h_der = round(h_izq - variacion, 2)
    elif preferencia == "Izquierdo":
        h_der = round(h_izq + variacion, 2)
    else: # Si es "Ambos", la variación es aleatoria (azar genético)
        direccion = random.choice([1, -1])
        h_der = round(h_izq + (variacion * direccion), 2)
    
    datos.append([nombre, preferencia, h_izq, h_der])

df = pd.DataFrame(datos, columns=[
    "Nombre", 
    "Preferencia de Hombro", 
    "Altura Hombro Izquierdo (cm)", 
    "Altura Hombro Derecho (cm)"
])

df.to_excel("costaleros_prueba.xlsx", index=False)
print("Excel generado con estadísticas médicas reales.")
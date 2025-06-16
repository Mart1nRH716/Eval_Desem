## Automatización de la carga de los reportes de excel de Evaluación y desempeño

Este asignación son dos script para cargar: la tabla de evaluación y desempeño para las UMAE´s (No se tenía previamnete estos datos cargados en la base de datos) y la automatización de la carga de EyD para las OOAD

### Características principales
* **Automatiza la carga de los reportes de excel de EyD en la base de datos tanto para OOAD como para CUMAE
* **Almacenamiento de archivos:** el programa realiza una búsqueda en las carpetas de "Reportes/OOAD o UMAE" Para obtener los reportes sin la necesidad de agregarlos al código.

### Requisitos del sistema
* **Pýthon**  >= 3.8
* **Dependencias:** Las dependencias adicionales se listan en el archivo `requirements.txt`.
* Nota: Para poder ejecutar los scripts, se deben de almacenar los resportes en la siguiente ruta: "Reportes/OOAD o UMAE". Y para los archivos de UMAE, se deben de renombrar los reportes de excel siguiento la siguiente estructura: "mes_año"

### Instalación
1. **Clonar el repositorio:**
   ```bash
   git clone 
   
## Instrucciones de Configuración
2. **Crear un entorno virtual
```bash
python -m venv venv
venv\Scripts\activate     # En Windows
source venv/bin/activate   # En Linux/Mac
```

## Instalar las dependencias
3. ** Ejecutar el siguiente comando:
```bash
pip install -r requirements.txt
```

Con estos pasos se pueden ejecutar los archivos .py



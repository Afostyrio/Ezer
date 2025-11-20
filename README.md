# Ezer
Un automatizador de la presentación de la premiación
## Instalación
Primero, hay que copiar el repositorio donde se encuentra Ezer:
```bash
git clone https://github.com/Afostyrio/Ezer.git
```
Ezer está diseñado para funcionar con Python 3, por lo que debe ser instalado en una computadora que cuente con Python 3. Recomiendo, además, usar un entorno virtual:
```bash
python -m venv venv
```
Se deben instalar los paquetes necesarios:
```bash
pip install os python-pptx pandas numpy
```
## Cómo utilizarlo
Ezer requiere varios inputs para funcionar, la mayoría se encuentran perfectamente organizados en la carpeta `inputs`:
- En la carpeta `inputs/csv` se requiere:
  1. Un archivo `Concursantes.csv` (la lista de todos los participantes) con una columna "Estado" y una columna "NOMBRE COMPLETO".
  2. Un archivo `Medallistas Individual.csv` (la lista de medallistas individuales). Las columnas deben ser (en ese orden): `Clave Estado,Estado,CLAVE,Nombre,A,Medalla,Nivel`. Un ejemplo de una fila de este archivo es: `GTO,Guanajuato,GTO_III_1,Joshua Sebastián González Torres,35,Plata,III`.
  3. Un archivo `Medallistas Equipos.csv` (la lista de medallistas por equipos). Las columnas deben ser `CLAVE,Estado,Medalla,Nivel` en ese orden. Por ejemplo: `GTO,Guanajuato,Oro,III`.
- En la carpeta `inputs/img/Individual`, se colocan todas las fotos individuales de los participantes. Las fotos deben seguir el esquema: `<CLAVE ESTADO>_<NIVEL>_<NUM>` por ejemplo, `GTO_III_1`. Ezer acepta PNG, JPEG y JPG.
- En la carpeta `inputs/img/Teams`, se colocan todas las fotos de los logos estatales. Las fotos deben seguir el esquema: `<CLAVE ESTADO>` por ejemplo, `GTO`. Ezer acepta PNG, JPEG y JPG.
- Un archivo `Plantilla.pptx`, el cual no debe ser modificado, por favor.
### El archivo `config`
El archivo `config` es el mero mole de Ezer. Este archivo determina el orden y el contenido de las diapositivas que se colocan a traves de comandos. Consideremos el archivo siguiente:
```
title
person:
+ title: Maestra de ceremonias
+ name: M. Isabel de Montserrat Avila Olivo
+ role: Delegada de Aguascalientes
+ image: inputs/img/Isabel.png
moment:
+ name: Bienvenida
parade
individual:
+ level: II
+ medal: Plata
team:
+ level: II
``` 
Desglosemos su estructura.

`title`
: Coloca la diapositiva de título "CEREMONIA DE PREMIACIÓN"

`person`
: Se utiliza para colocar una diapositiva "de persona" en la presentación. Se incluyen los siguientes parámetros
  - `title`: El título de la diapositiva ("Bienvenida", "Palabras de un alumno", "Un discurso repetido").
  - `name`: el nombre de la persona.
  - `role`: el título o rol de la persona ("Presidenta de la OMM", "Delegada de Aguascalientes", "Director del comité académico de la OMMEB", "Emisario de los Gorgonitas").
  - `image`: la ubicación y nombre de la foto de la persona.

`moment`
: Se utiliza para marcar secciones de una sola diapositiva en la presentación.
  - `name`: El nombre del "momento": "Palabras de los patrocinadores", "Agradecimientos a la sede".

`parade`
: Inserta el desfile de delegaciones (utiliza el archivo `inputs/csv/Concursantes.csv`).

`individual`
: Inserta las medallas individuales de acuerdo a los siguientes parámetros. Utiliza el archivo `inputs/csv/Medallistas Individual.csv`.
  - `level`: el nivel en números romanos (`I`, `II`, `III`).
  - `medal`: la medalla que se está premiando (`Oro`, `Plata`, `Bronce`, `Mención Honorífica`)

`team`
: Inserta las medallas por equipos. Utiliza el archivo `inputs/csv/Medallistas Equipos.csv`. Este comando considera empates en lugares.
  - `level`: el nivel en números romanos (`I`, `II`, `III`).

# Configuración de la institución

`institucion.json` guarda los datos del colegio que usan **todos** los prompts: nombre, ciudad, lema, rutas del escudo y membrete de los documentos (título, código de formato, versión y fecha).

- El generador de pruebas lo busca solo (o con `--institucion otro.json`).
- Las rutas de `escudo` son relativas a este archivo.
- Para usar el repositorio con otra institución, edita este archivo y cambia las imágenes con `herramientas/preparar-escudo.py`.
- Un `config.membrete` dentro de un `exam.json` tiene prioridad sobre este archivo (útil para un formato especial de una sola prueba).

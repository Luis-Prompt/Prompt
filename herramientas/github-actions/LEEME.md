# Verificación automática (pendiente de activar)

`verificar-examen.yml` prueba, cada vez que cambia el prompt de examen, que el ejemplo sigue generando los dos Word.

GitHub solo lo ejecuta si el archivo está en `.github/workflows/`. Para activarlo, cualquiera de estas dos vías:

1. **Desde la web:** en el repositorio, *Add file → Create new file*, nombre `.github/workflows/verificar-examen.yml`, pega el contenido de este archivo y guarda.
2. **Con Claude:** crea un token nuevo con el permiso **Workflows: Read and write** y pídele que mueva el archivo.

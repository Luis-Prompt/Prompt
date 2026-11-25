# Seguridad

## Qué NO se debe subir nunca
- Claves, tokens (GitHub, Anthropic, OpenAI…), contraseñas o archivos `.env`.
- Datos personales de estudiantes (nombres con notas, documentos de identidad, fotos).
- Datos personales de docentes más allá del nombre profesional (el `perfil_docente.json` real **no** se sube; solo el ejemplo).

`npm run verificar:calidad` revisa automáticamente patrones de tokens y llaves privadas, y el CI lo ejecuta en cada cambio.

## Si se filtra un token
1. **Revócalo de inmediato** en GitHub → Settings → Developer settings → Personal access tokens.
2. Crea uno nuevo con el mínimo de permisos (solo este repositorio; *Contents* y, si hace falta, *Workflows*) y fecha de vencimiento corta.
3. Nunca lo pegues en chats, issues ni commits.

## Reportar un problema de seguridad
Abre un aviso privado en la pestaña **Security → Report a vulnerability** del repositorio. No publiques los detalles en un Issue abierto.

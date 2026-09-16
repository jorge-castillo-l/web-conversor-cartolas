# Conversor de cartolas para Skualo (web)

Página estática para GitHub Pages. Pasa la cartola de **BCI**, **Banco Estado**, **Itaú** (`.xls` o `.xlsx`) y **Santander** al Excel que Skualo acepta en Tesorería → Cartolas.

Los archivos se leen **en el navegador**. No hay servidor ni API: no se suben cartolas a internet.

La página usa los logotipos públicos de Skualo, BCI, Banco Estado, Itaú y Santander solo para reconocer el origen y el destino. Las marcas son de sus dueños; el conversor no está afiliado. Fuentes en `img/ATRIBUCIONES.txt`.

Esta carpeta es independiente de `conversor-cartolas/` (el programa de escritorio).

## Cómo publicarla en GitHub Pages

1. Cree un repositorio **solo con esta carpeta** (no suba cartolas de clientes).
2. En GitHub: Settings → Pages → Deploy from a branch → `main` / `/ (root)`.
3. El enlace queda `https://SU_USUARIO.github.io/NOMBRE_DEL_REPO/`.
4. Si ya tiene una Page de programador, copie estos archivos a una subcarpeta (`conversor-cartolas/`) y enlace desde el portfolio.

## Cómo probarla en el PC

Abra `index.html` en el navegador, o desde esta carpeta:

```
python -m http.server 8080
```

Luego entre a `http://localhost:8080`.

## Qué no va aquí

- RUT ni nombres de una empresa cliente
- Cartolas de ejemplo con datos reales
- El `conversor.exe` de escritorio

SheetJS `xlsx` 0.18.5 (Apache-2.0) está en `vendor/` para leer `.xls` de Itaú sin instalar nada.

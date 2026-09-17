# Conversor de cartolas para Skualo

Aplicación web estática que convierte la cartola de **BCI**, **Banco Estado**, **Itaú** (`.xls` o `.xlsx`) y **Santander** al Excel que Skualo acepta en Tesorería → Cartolas.

Los archivos se leen **en el navegador**. No hay servidor ni API: las cartolas no se suben a internet.

La página usa los logotipos públicos de Skualo, BCI, Banco Estado, Itaú y Santander solo para identificar origen y destino. Las marcas son de sus dueños; el conversor no está afiliado. Fuentes en `img/ATRIBUCIONES.txt`.

## Cómo probarla

Abra `index.html` en el navegador, o desde esta carpeta:

```
python -m http.server 8080
```

Luego entre a `http://localhost:8080`.

SheetJS `xlsx` 0.18.5 (Apache-2.0) está en `vendor/` para leer `.xls` de Itaú sin instalar nada.

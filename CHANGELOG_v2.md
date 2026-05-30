# 📸 WEBAPP v2.0 - Conversión Automática de Formatos

## ✨ Nuevas Features

### Auto-conversión a JPG
La webapp ahora detecta automáticamente el formato de imagen y lo convierte a JPG:

```
Usuario sube:          Webapp convierte a:     Procesa:
├─ HEIC (iPhone)   →  JPG (95% quality)   →  Claude Vision
├─ PNG              →  JPG                →  Claude Vision
├─ WEBP             →  JPG                →  Claude Vision
├─ BMP              →  JPG                →  Claude Vision
├─ TIFF             →  JPG                →  Claude Vision
├─ GIF              →  JPG                →  Claude Vision
└─ JPG (nativo)     →  JPG (sin cambios)  →  Claude Vision
```

---

## 🔧 Cambios técnicos

### Nueva función: `convertir_a_jpg()`
```python
def convertir_a_jpg(image_bytes, filename):
    """
    - Detecta formato de imagen
    - Convierte a JPG si es necesario
    - Maneja transparencia (RGBA → RGB)
    - Retorna bytes optimizados (quality=95)
    """
```

### Flujo de procesamiento
1. Usuario sube imagen (cualquier formato)
2. `convertir_a_jpg()` detecta y convierte si es necesario
3. Se muestra el formato original en la UI: "(HEIC→JPG)" o "(JPG nativo)"
4. Se envía a Claude Vision
5. Se guardan datos + formato original en Excel

---

## 📊 Resultado en Excel

Ahora el Excel incluye una columna nueva `formato_original`:

| # | Comercio | ... | Archivo | Formato Original | Notas |
|---|----------|-----|---------|------------------|-------|
| 1 | SAM'S | ... | IMG_001.heic | HEIC | ✅ |
| 2 | HEB | ... | IMG_002.png | PNG | ✅ |
| 3 | OXXO | ... | IMG_003.jpg | JPG | ✅ |

---

## 🎯 Ventajas

✅ **Usuario no tiene que preconvertir**  
✅ **Maneja cámaras modernas (HEIC)**  
✅ **Compatible con cualquier dispositivo**  
✅ **Transparencia automática (PNG)**  
✅ **Compresión optimizada (quality=95)**  
✅ **Rápido (todo en memoria)**  

---

## 📦 Requisitos

Ya está incluido en `requirements.txt`:
```
Pillow==11.3.0  ← Para conversión de imágenes
```

---

## 🚀 Deploy

El deploy es igual:
1. Sube a GitHub
2. Conecta Streamlit Cloud
3. Agrega `ANTHROPIC_API_KEY` en secrets
4. ¡Listo!

---

## 💡 Ejemplo de uso

```
Ejecutivo de Grupo Barreda:
"Tengo recibos de transportes tomados con mi iPhone (HEIC)"
    ↓
Sube 30 fotos HEIC a la webapp
    ↓
Webapp detecta: "HEIC → JPG"
    ↓
Procesa con Claude Vision
    ↓
Descarga Excel con datos listos
```

---

## 🔍 Ver convertidor en acción

Cuando ejecutas la webapp:
```
Procesando 1/47: IMG_4608.heic (HEIC→JPG)
Procesando 2/47: IMG_4609.png (PNG→JPG)
Procesando 3/47: IMG_4610.jpg (JPG nativo)
...
```

El paréntesis muestra la conversión realizada.

---

## ¿Problemas?

Si una imagen no se puede convertir:
- Se reporta en la sección "Imágenes con error"
- Se continúa con las siguientes
- Usuario ve exactamente cuál falló y por qué

---

**Actualizado:** Mayo 30, 2026  
**Versión:** 2.0  
**Estatus:** ✅ Listo para producción

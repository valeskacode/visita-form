# Visita a Clientes de Pequeña Empresa — v2 (mobile-first)

App de campo en Streamlit, rediseñada con enfoque 100% celular (también
funciona en PC), siguiendo el flujo:

**Búsqueda y carga → Evaluación de crédito (criterios) → Ficha del cliente
→ Ingresos y gastos → Ubicación (visita con foto + GPS) → Reporte (Word y PDF)**

## Qué trae esta versión

- **Lee la hoja `MUESTRA_FINAL`** del Excel directamente (si no existe, usa
  la primera hoja y te avisa).
- **Carga rápida en PC y celular por igual**: el archivo se procesa en el
  servidor (no en tu dispositivo) y se cachea, así que no se vuelve a leer
  en cada clic.
- **Criterios para la visita** (la tabla que enviaste) como paso de
  evaluación, con panel que se pone en rojo cuando hay algo marcado.
- **Ficha de cliente** de solo lectura con DNI, saldo, importe, % pagado,
  badge de riesgo según días de atraso.
- **Ingresos y gastos** con resultado neto y margen calculados al vuelo.
- **Visitas** (domicilio / negocio / aval) cada una con foto (cámara o
  galería), captura de GPS, fecha/hora y comentarios.
- **Reporte final en Word (.docx) Y en PDF**, ambos descargables.
- **Recuperación de avance**: si cierras la app a medias, al volver a
  buscar al mismo cliente con el mismo usuario te pregunta si quieres
  continuar donde quedaste o empezar de nuevo.
- **Historial**: cada vez que generas un Word o un PDF se anota quién lo
  generó, cuándo, y con qué criterios — visible en la pestaña de reporte.

## Lo que NO incluye (para que no lo esperes y te confunda)

- **"Escanear DNI" con la cámara no está implementado.** Hacerlo bien
  requiere un servicio de OCR (lectura de texto en imágenes), que no se
  agregó en esta versión para no prometer algo poco confiable. La búsqueda
  por DNI escribiéndolo a mano sí funciona normal.
- El stepper de 5 pasos de tu mockup de "Nueva Visita" se simplificó a
  4 bloques dentro de una sola pantalla por cada tipo de visita (foto, GPS,
  datos del lugar, observaciones) en vez de pantallas separadas paso a
  paso — mismo contenido, con menos clics. Si prefieres el wizard paso a
  paso real, se puede armar después.
- Las secciones de **Garantías** y **Deuda RCC** del formulario original en
  PDF no tienen pantalla propia en este flujo de 6 pasos (no se pidieron
  explícitamente). El generador de reportes ya las soporta internamente;
  si las quieres de vuelta como parte del flujo, se agregan fácil.

## Estructura del proyecto

```
app.py                      # router principal + las 6 pantallas
utils/
  helpers.py                 # lógica compartida: Excel, autoguardado,
                              # historial, generación de Word/PDF
assets/
  style.css                  # estilos mobile-first
.streamlit/
  config.toml                # tema (incluye base="light" para que no se
                              # vea negro en celulares con modo oscuro)
data/                         # se crea solo; guarda avances e historial
requirements.txt
.gitignore
```

## Importante sobre dónde vive la información

`data/drafts/` (avances en curso) y `data/historial_visitas.xlsx`
(historial de reportes generados) se guardan en el disco del servidor
donde corra la app:

- **Corriendo en tu PC o un servidor propio**: persiste mientras no borres
  la carpeta — sirve perfecto para el día a día.
- **En Streamlit Community Cloud (plan gratis)**: el disco es temporal —
  se reinicia cuando la app se "duerme" por inactividad o se redespliega.
  Para uso diario intensivo en la nube gratuita, descarga el reporte de
  cada visita apenas se genere (el botón ya te lo recuerda), y no dependas
  del historial como respaldo único a largo plazo.

## Cómo subir esto a GitHub y publicarlo

Sigue exactamente los mismos pasos que ya armamos antes (crear repo, subir
archivos por el navegador, conectar en share.streamlit.io). Dos detalles
nuevos para esta versión:

1. Al subir los archivos a GitHub, asegúrate de subir también las carpetas
   **`utils/`** y **`assets/`** completas (con `helpers.py` y `style.css`
   adentro) — no solo `app.py`. Sin ellas la app no arranca.
2. En "Main file path" al desplegar en Streamlit Cloud, sigue siendo
   `app.py` (el router que vive en la raíz).

## Probar localmente antes de publicar

```bash
pip install -r requirements.txt
streamlit run app.py
```

Abre `http://localhost:8501` en tu navegador (en el mismo PC funciona
cámara/GPS por ser `localhost`; desde el celular en la misma red NO
funcionarán cámara/GPS por el tema de HTTPS que ya vimos — para eso,
publica en Streamlit Cloud).

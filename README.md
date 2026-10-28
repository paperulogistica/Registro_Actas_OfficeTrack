# Registro de Actas de Entrega/Devolución de Equipos

App web (Flask) que lee correos de Outlook (.msg) con Actas generadas por
OfficeTrack, o PDFs sueltos, extrae los datos de cada equipo y los acumula
en una base de datos (SQLite) — nunca se pierde lo ya procesado entre
sesiones. Cada fila queda con su fecha/hora de procesamiento puesta
automáticamente por el servidor. Desde la web se descarga en cualquier
momento el Excel completo, acumulado, con el mismo diseño ya validado.

Principios (igual que el proyecto de Factura+Guía+OC): 100% código
abierto, sin SharePoint/OneDrive, sin suscripciones de pago. Los archivos
originales (.msg/.pdf) NO se guardan en el servidor — solo se procesan en
memoria; lo único que se guarda es la fila de datos extraída.

## Estructura

```
webapp/
  app.py                     Rutas Flask (/, /subir, /descargar)
  requirements.txt
  Procfile                   Para desplegar en Render (gunicorn)
  templates/index.html       Página única: subir, ver preview, descargar
  registro_actas/
    parser.py                Lee .msg/.pdf y devuelve filas (misma lógica validada)
    db.py                     Base de datos acumulativa SQLite
    excel_export.py          Genera el Excel con el diseño de la maqueta
    data/                     Carpeta donde vive actas.db (se crea sola)
```

## Correr en local

```bash
cd webapp
python3 -m venv venv && source venv/bin/activate
pip install -r requirements.txt
python3 app.py
```

Abre http://localhost:5000

## Desplegar en Render (plan gratuito)

1. Sube esta carpeta a un repositorio de GitHub
   (ej. `paperulogistica/registro-actas`).
2. En Render: **New > Web Service**, conecta el repo.
3. Render detecta el `Procfile` automáticamente. Si pide configurarlo a mano:
   - Build command: `pip install -r requirements.txt`
   - Start command: `gunicorn app:app`
4. **Importante — el disco de Render es efímero** en el plan gratuito: si el
   servicio se reinicia o duerme por inactividad, se puede perder la base
   `actas.db` junto con todo lo acumulado. Para que la acumulación sea
   confiable en producción, cuando quieras pasar esto a "en serio" conviene:
   - Activar un **Persistent Disk** de Render (tiene costo, aunque bajo) y
     montar `registro_actas/data/` ahí, o
   - Cambiar `db.py` para usar una base de datos gestionada gratuita
     (ej. el plan free de Render Postgres, o Turso/LibSQL) en vez de un
     archivo SQLite local.
   Mientras se valida el sistema, el SQLite local funciona bien — solo hay
   que tener en cuenta que un reinicio del servicio puede reiniciar los
   datos si no hay disco persistente.
5. Variable de entorno opcional `SECRET_KEY` (para las notificaciones de
   Flask) — Render genera una por defecto si no la pones.

## Qué hace y qué no hace todavía

- Reconoce los formularios **TI-F-001** (celular) y **TI-F-002** (equipo de
  cómputo). Encontramos un tercer formato, **MA-F-028** ("Acta de Entrega o
  Devolución de Activos", para activos generales como sillas/muebles) que
  tiene una estructura distinta (sin datos de persona/DNI, con Placa/Estado
  en vez de IMEI) — por ahora se marca como "no reconocido" y queda
  disponible para completar a mano, en vez de inventar cómo llenarlo.
- Si quieres que MA-F-028 también se automatice, hay que darle su propia
  función de extracción en `parser.py` (su layout es distinto, no comparte
  columnas con TI-F-001/002).
- Un archivo .msg/.pdf que ya se procesó antes (se detecta por hash del
  contenido) no se vuelve a duplicar si se sube dos veces.
- El Excel se regenera desde la base de datos cada vez que se descarga
  (no se guarda un .xlsx fijo en el servidor).

# Generador de reportes de ventas

Convierte un CSV de ventas en un reporte de Excel listo para entregar: hoja de **detalle** y hoja de **resumen por producto** con totales y porcentaje de participación.

Pensado para pequeños negocios que hoy arman estos reportes a mano en Excel.

## Instalación

```bash
git clone <URL-de-tu-repositorio>
cd reportes-python
python3 -m venv venv && source venv/bin/activate
pip install -r requirements.txt
```

## Uso

```bash
python generar_reporte.py datos_ejemplo/ventas.csv reporte.xlsx
```

Si omites el nombre de salida, se crea `reporte.xlsx`.

## Formato de entrada

CSV con las columnas `fecha, producto, cantidad, precio`:

| fecha      | producto | cantidad | precio |
|------------|----------|----------|--------|
| 2026-10-01 | Arroz    | 10       | 4500   |
| 2026-10-01 | Aceite   | 5        | 12000  |

- Separador `,` o `;` (se detecta solo). Con `;` los decimales van con coma, como lo guarda Excel en español.
- Fechas en `AAAA-MM-DD` o `DD/MM/AAAA`.
- Los encabezados no distinguen mayúsculas ni espacios sobrantes.

## Validaciones

El script se detiene con un mensaje claro, sin trazas técnicas, si encuentra:

- archivo inexistente o vacío
- columnas faltantes
- filas con fechas o números inválidos (indica el número de fila)
- cantidades o precios negativos

## Pruebas

```bash
pip install -r requirements-dev.txt
pytest
```

## Limitaciones conocidas

- Solo lee CSV (no `.xlsx` de entrada).
- La salida es Excel; todavía no genera PDF ni gráficos.

## Licencia

MIT. Consulta [`LICENSE`](LICENSE).
